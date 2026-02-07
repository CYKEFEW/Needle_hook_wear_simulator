# -*- coding: utf-8 -*-
"""
needle_hook_wear_simulator_gui.py

针钩磨损平台检测全过程仿真（核心引擎）
- UTF-8 编码 / 中文注释
- fs 与 duration 仅用于生成时间轴（输出点数），不做更细步长采样过程仿真
- 机械周期扰动主频仅由转速输入：f_mech = rpm/60*m（默认300rpm）
- 陷波滤波器 Q 不手动输入：根据 rpm 自动估算（Q≈clamp(15,80,rpm/10)）
- 稳定段基线记为 μss；超限阈值记为 μth；tlife 数值仅在图例显示
- 规则：第一次超限后不再判定稳定段窗口（稳定段并集仅来自超限前）

输出：
- xlsx（同一xlsx可多工作表，必要时自动拆分）
- png 曲线图（可选中文/英文）
- summary.json

主要接口：
- simulate(cfg, seed, progress_cb) -> res（仅生成仿真结果，不写文件）
- export_xlsx(res, out_dir, progress_cb) -> xlsx_path
- export_plots(res, out_dir, lang, progress_cb) -> dict(paths)
"""

import os
import json
import math
from dataclasses import dataclass, asdict
from typing import Optional, Dict, Any, Callable, List, Tuple

import numpy as np
import pandas as pd
import matplotlib
import matplotlib.pyplot as plt
from matplotlib import font_manager
from openpyxl import Workbook

# SciPy 可选：更标准的 IIR 滤波
try:
    from scipy import signal  # type: ignore
    HAVE_SCIPY = True
except Exception:
    HAVE_SCIPY = False

ProgressCB = Optional[Callable[[float, str], None]]  # cb(pct 0~100, msg)


def _ensure_dir(p: str) -> None:
    os.makedirs(p, exist_ok=True)


def _cb(cb: ProgressCB, pct: float, msg: str) -> None:
    if cb:
        cb(float(max(0.0, min(100.0, pct))), str(msg))


def _notch_q_from_rpm(rpm: float) -> float:
    """根据转速估计陷波 Q：Q = clamp(15,80,rpm/10)"""
    try:
        rpm = float(rpm)
    except Exception:
        rpm = 300.0
    return float(max(15.0, min(80.0, rpm / 10.0)))

def _notch_q_from_rpm_vec(rpm_t: np.ndarray) -> np.ndarray:
    """向量化计算陷波 Q（与 rpm 对齐）"""
    arr = np.asarray(rpm_t, dtype=float)
    q = arr / 10.0
    q = np.clip(q, 15.0, 80.0)
    q[~np.isfinite(arr)] = 15.0
    return q.astype(float)


def _setup_chinese_font(cb: ProgressCB = None, pct: float = 2.0) -> Dict[str, Any]:
    """尝试设置 matplotlib 中文字体，避免导出图片中文乱码"""
    candidates = [
        "Microsoft YaHei", "微软雅黑",
        "SimHei", "黑体",
        "Noto Sans CJK SC", "Noto Sans CJK",
        "WenQuanYi Micro Hei", "文泉驿微米黑",
        "PingFang SC", "Heiti SC",
        "Source Han Sans SC", "思源黑体",
        "Arial Unicode MS",
    ]
    available = {f.name for f in font_manager.fontManager.ttflist}
    chosen = None
    for c in candidates:
        if c in available:
            chosen = c
            break

    matplotlib.rcParams["axes.unicode_minus"] = False
    if chosen:
        matplotlib.rcParams["font.sans-serif"] = [chosen]
        _cb(cb, pct, f"已设置绘图中文字体：{chosen}")
        return {"font_ok": True, "font_name": chosen}

    _cb(cb, pct, "未检测到常见中文字体：中文图可能仍会乱码（可安装微软雅黑/黑体/Noto Sans CJK等）")
    return {"font_ok": False, "font_name": None}


def setup_plot_font(lang: str = "zh", progress_cb: ProgressCB = None, pct: float = 2.0) -> Dict[str, Any]:
    """按语言准备绘图字体（用于 GUI 预处理或导出阶段）"""
    if str(lang).lower().startswith("zh"):
        return _setup_chinese_font(progress_cb, pct=pct)
    matplotlib.rcParams["axes.unicode_minus"] = False
    return {"font_ok": True, "font_name": None}


def _downsample_for_plot(x: np.ndarray, y: np.ndarray, max_points: int = 200_000):
    """绘图降采样：只影响绘图，不影响导出数据"""
    n = len(x)
    if n <= max_points:
        return x, y
    step = max(1, n // max_points)
    return x[::step], y[::step]


def _moving_average(x: np.ndarray, win: int) -> np.ndarray:
    """移动平均（无 SciPy 时用作低通近似）"""
    if win <= 1:
        return x.copy()
    win = int(win)
    kernel = np.ones(win, dtype=float) / float(win)
    pad = win // 2
    xpad = np.pad(x, (pad, pad), mode="reflect")
    return np.convolve(xpad, kernel, mode="valid")


def _hampel_filter_nan(x: np.ndarray, win: int, n_sigmas: float = 3.0) -> np.ndarray:
    """Hampel：滚动中位数+MAD 检测异常点，异常点置 NaN"""
    if win <= 1:
        return x.copy()
    s = pd.Series(x)
    med = s.rolling(win, center=True, min_periods=max(3, win // 2)).median()
    mad = (s - med).abs().rolling(win, center=True, min_periods=max(3, win // 2)).median()
    sigma = 1.4826 * mad
    outlier = (s - med).abs() > (n_sigmas * sigma)
    y = s.astype(float).copy()
    y[outlier] = np.nan
    return y.to_numpy()


def _notch_freq_domain(x: np.ndarray, fs: float, f0: float, q: float) -> np.ndarray:
    """无 SciPy：频域抑制（用于仿真展示足够）"""
    if f0 <= 0 or f0 >= fs / 2:
        return x.copy()
    X = np.fft.rfft(x)
    freqs = np.fft.rfftfreq(len(x), d=1.0 / fs)
    bw = max(0.05, f0 / max(1.0, q))
    mask = (freqs > (f0 - bw)) & (freqs < (f0 + bw))
    X[mask] = 0
    return np.fft.irfft(X, n=len(x))


def _iir_notch_time_varying(
    x: np.ndarray,
    fs: float,
    f0_t: np.ndarray,
    q_t: np.ndarray,
    block_s: float = 10.0,
) -> np.ndarray:
    """随时间变化的陷波：分块 + 重叠加权近似"""
    x = np.asarray(x, dtype=float)
    n = len(x)
    if n == 0:
        return x.copy()

    f0_t = np.asarray(f0_t, dtype=float)
    q_t = np.asarray(q_t, dtype=float)
    if len(f0_t) != n:
        f0_t = np.full(n, float(np.nanmedian(f0_t)) if len(f0_t) > 0 else 0.0)
    if len(q_t) != n:
        q_t = np.full(n, float(np.nanmedian(q_t)) if len(q_t) > 0 else 15.0)

    block = int(max(32, round(block_s * fs)))
    block = min(block, n)
    hop = max(1, block // 2)

    out = np.zeros_like(x)
    weight = np.zeros_like(x)

    for start in range(0, n, hop):
        end = min(n, start + block)
        seg = x[start:end]
        f_seg = f0_t[start:end]
        q_seg = q_t[start:end]
        m = np.isfinite(f_seg) & np.isfinite(q_seg)
        if not np.any(m):
            seg_f = seg
        else:
            f0 = float(np.median(f_seg[m]))
            qv = float(np.median(q_seg[m]))
            if f0 <= 0 or f0 >= fs / 2:
                seg_f = seg
            else:
                if HAVE_SCIPY:
                    b, a = signal.iirnotch(w0=f0, Q=qv, fs=fs)
                    try:
                        seg_f = signal.filtfilt(b, a, seg, method="gust")
                    except Exception:
                        try:
                            seg_f = signal.filtfilt(b, a, seg)
                        except Exception:
                            seg_f = signal.lfilter(b, a, seg)
                else:
                    seg_f = _notch_freq_domain(seg, fs, f0, qv)

        L = end - start
        if L <= 1:
            w = np.ones(L, dtype=float)
        else:
            w = np.hanning(L).astype(float)
        out[start:end] += seg_f * w
        weight[start:end] += w

    y = x.copy()
    m = weight > 1e-12
    y[m] = out[m] / weight[m]
    return y


def _iir_notch(x: np.ndarray, fs: float, f0: float | np.ndarray, q: float | np.ndarray, block_s: float = 10.0) -> np.ndarray:
    """陷波滤波：支持 f0/Q 随时间变化"""
    if np.ndim(f0) == 0 and np.ndim(q) == 0:
        f0 = float(f0)
        q = float(q)
        if f0 <= 0 or f0 >= fs / 2:
            return x.copy()
        if HAVE_SCIPY:
            b, a = signal.iirnotch(w0=f0, Q=q, fs=fs)
            try:
                return signal.filtfilt(b, a, x, method="gust")
            except Exception:
                return signal.filtfilt(b, a, x)
        return _notch_freq_domain(x, fs, f0, q)

    return _iir_notch_time_varying(x, fs, np.asarray(f0), np.asarray(q), block_s=block_s)


def _lowpass(x: np.ndarray, fs: float, fc: float, order: int = 3) -> np.ndarray:
    """低通滤波：优先 SciPy butter+filtfilt，否则移动平均近似"""
    if fc <= 0:
        return x.copy()

    if HAVE_SCIPY:
        wn = min(0.99, fc / (fs / 2))
        b, a = signal.butter(order, wn, btype="low")
        try:
            return signal.filtfilt(b, a, x, method="gust")
        except Exception:
            return signal.filtfilt(b, a, x)

    win = int(max(1, round(fs / max(1e-6, fc))))
    win = min(win, 2001)
    return _moving_average(x, win)


def _first_exceed_index(mu_f: np.ndarray, thr: float):
    """第一次超限点索引（忽略NaN），若无则返回 None"""
    if thr is None or not np.isfinite(thr):
        return None
    m = np.isfinite(mu_f) & (mu_f > thr)
    if not np.any(m):
        return None
    return int(np.argmax(m))




def _first_exceed_run_start(mu_f: np.ndarray, thr: float, hold_count: int):
    """
    找到第一次“连续超限 hold_count 个点”的起点索引（忽略NaN），若无则返回 None。
    用途：用于“超限后不再判定稳定段”，但避免单点尖峰导致过早停止。
    """
    if thr is None or not np.isfinite(thr):
        return None
    hold_count = int(max(1, hold_count))
    count = 0
    for i, v in enumerate(mu_f):
        if np.isfinite(v) and v > thr:
            count += 1
            if count >= hold_count:
                return int(i - hold_count + 1)
        else:
            count = 0
    return None
def _write_xlsx_multisheet(
    xlsx_path: str,
    frames: Dict[str, pd.DataFrame],
    row_limit: int = 1_048_000,
    progress_cb: ProgressCB = None,
    base_pct: float = 0.0,
    span_pct: float = 100.0,
) -> None:
    """
    写入 xlsx（同一文件多工作表，必要时自动拆分）

    设计要点：
    - 直接使用 xlsxwriter 写入（比 openpyxl / pandas.to_excel 通常更快）
    - 若行数超过单sheet上限（约 1048576），会自动在同一个 xlsx 中拆分多个工作表
    """
    import xlsxwriter

    total_rows = sum(len(df) for df in frames.values())
    total_rows = max(1, total_rows)
    written = 0
    total_cells = int(sum(int(df.shape[0]) * int(df.shape[1]) for df in frames.values()))

    # 小规模数据按列写更快；大规模数据使用常量内存按行写更稳
    max_cells_for_column = 5_000_000
    use_column_write = total_cells <= max_cells_for_column

    def _display_width(s: str) -> int:
        import unicodedata
        w = 0
        for ch in s:
            if unicodedata.east_asian_width(ch) in ("W", "F"):
                w += 2
            else:
                w += 1
        return w

    def _cell_len(v) -> int:
        if v is None:
            return 0
        try:
            if isinstance(v, float) and not np.isfinite(v):
                return 0
        except Exception:
            pass
        return _display_width(str(v))

    def _normalize_value(v):
        if v is None:
            return None
        try:
            if isinstance(v, float) and not np.isfinite(v):
                return None
        except Exception:
            pass
        return v

    def _normalize_row(row):
        return [_normalize_value(v) for v in row]

    def _normalize_column(data):
        if not data:
            return data
        try:
            arr = np.asarray(data, dtype=float)
            if np.isfinite(arr).all():
                return data
        except Exception:
            pass
        return [_normalize_value(v) for v in data]

    def _col_width(col_name: str, sample_vals: List[Any]) -> float:
        max_len = _cell_len(col_name)
        for v in sample_vals:
            max_len = max(max_len, _cell_len(v))
        # 适度留白，限制最大宽度，避免异常值导致过宽
        width = max(8, min(40, max_len + 2))
        return float(width)

    def _p(msg: str):
        pct = base_pct + span_pct * (written / total_rows)
        _cb(progress_cb, min(99.0, pct), msg)

    wb = xlsxwriter.Workbook(
        xlsx_path,
        {"constant_memory": not use_column_write, "nan_inf_to_errors": True},
    )
    try:
        for name, df in frames.items():
            n = len(df)
            if n <= row_limit:
                parts = [(name, df)]
            else:
                num = int(math.ceil(n / row_limit))
                parts = []
                for i in range(num):
                    s = i * row_limit
                    e = min(n, (i + 1) * row_limit)
                    parts.append((f"{name}_{i + 1}", df.iloc[s:e]))

            for sheet, part_df in parts:
                ws = wb.add_worksheet(sheet[:31])  # Excel sheet name limit
                cols = list(part_df.columns)

                # 表头
                ws.write_row(0, 0, cols)

                if use_column_write:
                    # 按列写入（速度快，但会占用更多内存）
                    for c, col in enumerate(cols):
                        data = _normalize_column(part_df[col].tolist())
                        ws.write_column(1, c, data)
                else:
                    # 常量内存模式下必须按行写入，否则会被丢弃
                    row_idx = 1
                    for row in part_df.itertuples(index=False, name=None):
                        ws.write_row(row_idx, 0, _normalize_row(row))
                        row_idx += 1

                # 根据前两行（表头 + 前两条数据）自动调整列宽
                sample_df = part_df.head(2)
                for c, col in enumerate(cols):
                    sample_vals = []
                    if not sample_df.empty:
                        sample_vals = sample_df[col].tolist()
                    ws.set_column(c, c, _col_width(col, sample_vals))

                written += len(part_df)
                _p(f"写入 xlsx：{sheet}（{written}/{total_rows} 行）")
    finally:
        wb.close()

    _cb(progress_cb, min(99.0, base_pct + span_pct), "xlsx 写入结束")

@dataclass
class SimConfig:
    # 核心输入：fs 与 duration 仅用于生成时间轴
    theta_deg: float = 100.0
    t_set_N: float = 0.5
    fs_Hz: float = 50.0
    duration_s: float = 36000.0

    # 机械扰动：仅由 rpm 换算
    # 兼容字段：rpm（固定值），推荐用 rpm_min/rpm_max + tau_rpm_s 做慢变
    rpm: float = 300.0
    rpm_min: float = 285.0
    rpm_max: float = 315.0
    tau_rpm_s: float = 1200.0
    mech_harmonic: int = 1

    # 扰动（开环明显，闭环衰减）
    noise_rms_open: float = 0.0275
    noise_rms_closed: float = 0.008
    mech_amp_open: float = 0.09
    mech_amp_closed: float = 0.02
    drift_amp_open: float = 0.055
    drift_amp_closed: float = 0.02
    drift_freq_hz: float = 0.000175

    # 阶段时间比例（三段之和会在 validate() 中归一化为 1）
    # - GUI 使用滑块直接调这三个比例
    # - 兼容字段：runin_ratio / severe_start_ratio 会由这里自动推导
    phase_runin_ratio: float = 0.12
    phase_stable_ratio: float = 0.70
    phase_severe_ratio: float = 0.18

    # 三阶段摩擦系数范围（用于生成 μ(t) 真值；同一 seed 下可重复）
    # - 若 min==max，则等效为固定值
    mu_runin_min: float = 0.30
    mu_runin_max: float = 0.40
    mu_stable_min: float = 0.22
    mu_stable_max: float = 0.28
    mu_severe_min: float = 0.35
    mu_severe_max: float = 0.60


    # 阶段过渡速度系数（用于调整“段间过渡速度”，保证连续）
    # - 磨合→稳定：系数越大，衰减越快（过渡越快）
    # - 稳定→加速：系数越大，S 形上升越陡（过渡越快）
    trans_runin2stable_k: float = 2.0
    trans_stable2severe_k: float = 5.0

    # 开环-闭环对比输出范围（单位：s）
    # - (0, -1) 表示全部输出（默认）
    compare_t_start_s: float = 0.0
    compare_t_end_s: float = -1.0

    # 兼容旧字段（内部会根据“范围”生成使用值，并写入 derived 输出）
    mu_runin_start: float = 0.35
    mu_stable: float = 0.25
    mu_severe_end: float = 0.55
    runin_ratio: float = 0.12
    severe_start_ratio: float = 0.82

    # 张力测量噪声
    sensor_rms: float = 0.005


    # ===== 扰动参数范围设置（用于准周期/不确定扰动；单位同原参数）=====
    # 说明：*_min/*_max 为范围；若只需固定值，可令 min=max。

    # 高频噪声强度（RMS，N）
    noise_rms_open_min: float = 0.015
    noise_rms_open_max: float = 0.040
    noise_rms_closed_min: float = 0.004
    noise_rms_closed_max: float = 0.012

    # 机械周期扰动幅值（N）
    mech_amp_open_min: float = 0.06
    mech_amp_open_max: float = 0.12
    mech_amp_closed_min: float = 0.01
    mech_amp_closed_max: float = 0.03

    # 低频准周期漂移幅值（N）
    drift_amp_open_min: float = 0.03
    drift_amp_open_max: float = 0.08
    drift_amp_closed_min: float = 0.01
    drift_amp_closed_max: float = 0.03

    # 低频漂移频率范围（Hz）：准周期 -> 两频率分量叠加
    drift_freq_hz_min: float = 0.00005
    drift_freq_hz_max: float = 0.00030

    # 传感器测量噪声（RMS，N）
    sensor_rms_min: float = 0.002
    sensor_rms_max: float = 0.008

    # 慢变时间常数（s）：用于让扰动“范围随时间慢变”（更接近真实准周期/工况漂移）
    # - <=0 表示不启用慢变：仍按“每次仿真从范围内抽一次常量”的旧逻辑（兼容）
    # - 建议：机械周期扰动/噪声RMS 的 τ 取 300~3000s；漂移相关 τ 取 1800~20000s
    tau_mech_s: float = 1200.0
    tau_noise_s: float = 300.0
    tau_sensor_s: float = 300.0
    tau_drift_amp_s: float = 900.0
    tau_drift_freq_s: float = 4800.0

    # 门控/裁剪（抑制“比值+对数”放大）
    tmin_gate_N: float = 0.08
    ratio_clip_min: float = 0.3
    ratio_clip_max: float = 8.0

    # 滤波
    hampel_win_s: float = 1.0
    hampel_nsig: float = 3.0
    lowpass_fc_hz: float = 2.5

    # 稳定段/寿命判据
    stable_win_s: float = 1200.0
    # 最短连续稳定段时长 Whold（s）：用于判定“连续稳定段”是否成立
    # - 若连续稳定窗口并集覆盖时间 < Whold，则该段不计为稳定段，也不用于 μss 计算
    stable_hold_s: float = 7200.0

    stable_sigma_max: float = 0.03  # 默认 0.03
    stable_slope_max: float = 0.015  # Δμ_max：稳定窗口内允许的最大总漂移量（默认按 Wss=300s 约等效 5e-5/s）
    stable_valid_min: float = 0.9
    fail_delta: float = 0.25
    fail_hold_s: float = 300.0

    # 导出/绘图
    export_stride: int = 1
    plot_max_points: int = 2_000_000

    def validate(self) -> None:
        assert self.fs_Hz > 0
        assert self.duration_s > 0
        assert self.export_stride >= 1
        assert 0 < self.theta_deg < 1080
        assert self.mech_harmonic >= 1
        # rpm 范围合法性
        try:
            rmin = float(getattr(self, "rpm_min", self.rpm))
            rmax = float(getattr(self, "rpm_max", self.rpm))
        except Exception:
            rmin = float(self.rpm)
            rmax = float(self.rpm)
        if rmin <= 0 or rmax <= 0:
            raise ValueError("rpm 必须>0（仅使用转速输入机械主频）")
        if rmin > rmax:
            rmin, rmax = rmax, rmin
        # 若 rpm_min/max 仍是默认值且 rpm 不同，则以 rpm 为准（兼容旧用法）
        try:
            rpm_val = float(self.rpm)
            if abs(rmin - rmax) < 1e-12 and abs(rmin - 300.0) < 1e-9 and abs(rpm_val - 300.0) > 1e-9:
                rmin = rpm_val
                rmax = rpm_val
        except Exception:
            pass
        self.rpm_min = float(rmin)
        self.rpm_max = float(rmax)
        # 同步 rpm 为当前范围中值（便于兼容旧逻辑）
        self.rpm = float(0.5 * (rmin + rmax))
        # 转速慢变时间常数
        self.tau_rpm_s = float(max(0.0, getattr(self, "tau_rpm_s", 0.0)))

        # 阶段比例归一化
        r1 = max(0.0, float(self.phase_runin_ratio))
        r2 = max(0.0, float(self.phase_stable_ratio))
        r3 = max(0.0, float(self.phase_severe_ratio))
        s = r1 + r2 + r3
        if s <= 1e-9:
            r1, r2, r3 = 0.12, 0.70, 0.18
            s = r1 + r2 + r3
        self.phase_runin_ratio = r1 / s
        self.phase_stable_ratio = r2 / s
        self.phase_severe_ratio = r3 / s

        # 兼容旧边界字段（用于旧逻辑/输出）
        self.runin_ratio = float(self.phase_runin_ratio)
        self.severe_start_ratio = float(self.phase_runin_ratio + self.phase_stable_ratio)
        self.severe_start_ratio = max(self.runin_ratio + 1e-6, min(0.999999, self.severe_start_ratio))

        # μ 范围合法性（min/max 自动交换）
        # 过渡速度系数范围保护
        self.trans_runin2stable_k = float(max(0.05, min(20.0, self.trans_runin2stable_k)))
        self.trans_stable2severe_k = float(max(0.05, min(20.0, self.trans_stable2severe_k)))

        # 开环-闭环对比输出范围保护
        self.compare_t_start_s = float(max(0.0, self.compare_t_start_s))
        if float(self.compare_t_end_s) != -1.0:
            self.compare_t_end_s = float(max(self.compare_t_start_s, self.compare_t_end_s))        # 连续稳定段最短时长（s）
        self.stable_hold_s = float(max(0.0, self.stable_hold_s))        # 兼容旧版本：若 stable_slope_max 非常小（<1e-3），通常表示“斜率阈值(1/s)”，自动换算为总漂移量阈值
        # Δμ_max ≈ g_max * Wss
        if float(self.stable_slope_max) < 1e-3:
            self.stable_slope_max = float(self.stable_slope_max) * float(self.stable_win_s)
        # 最小下限保护
        self.stable_slope_max = float(max(0.0, self.stable_slope_max))
        # 扰动范围参数：确保 min<=max，且不为负
        def _clamp_range(a: float, b: float, lo: float = 0.0, hi: float = 1e9):
            a = float(max(lo, min(hi, a)))
            b = float(max(lo, min(hi, b)))
            if a > b:
                a, b = b, a
            return a, b

        self.noise_rms_open_min, self.noise_rms_open_max = _clamp_range(getattr(self, "noise_rms_open_min", self.noise_rms_open),
                                                                         getattr(self, "noise_rms_open_max", self.noise_rms_open),
                                                                         lo=0.0)
        self.noise_rms_closed_min, self.noise_rms_closed_max = _clamp_range(getattr(self, "noise_rms_closed_min", self.noise_rms_closed),
                                                                             getattr(self, "noise_rms_closed_max", self.noise_rms_closed),
                                                                             lo=0.0)

        self.mech_amp_open_min, self.mech_amp_open_max = _clamp_range(getattr(self, "mech_amp_open_min", self.mech_amp_open),
                                                                      getattr(self, "mech_amp_open_max", self.mech_amp_open),
                                                                      lo=0.0)
        self.mech_amp_closed_min, self.mech_amp_closed_max = _clamp_range(getattr(self, "mech_amp_closed_min", self.mech_amp_closed),
                                                                          getattr(self, "mech_amp_closed_max", self.mech_amp_closed),
                                                                          lo=0.0)

        self.drift_amp_open_min, self.drift_amp_open_max = _clamp_range(getattr(self, "drift_amp_open_min", self.drift_amp_open),
                                                                        getattr(self, "drift_amp_open_max", self.drift_amp_open),
                                                                        lo=0.0)
        self.drift_amp_closed_min, self.drift_amp_closed_max = _clamp_range(getattr(self, "drift_amp_closed_min", self.drift_amp_closed),
                                                                            getattr(self, "drift_amp_closed_max", self.drift_amp_closed),
                                                                            lo=0.0)

        self.drift_freq_hz_min, self.drift_freq_hz_max = _clamp_range(getattr(self, "drift_freq_hz_min", self.drift_freq_hz),
                                                                      getattr(self, "drift_freq_hz_max", self.drift_freq_hz),
                                                                      lo=0.0, hi=1.0)

        self.sensor_rms_min, self.sensor_rms_max = _clamp_range(getattr(self, "sensor_rms_min", self.sensor_rms),
                                                                getattr(self, "sensor_rms_max", self.sensor_rms),
                                                                lo=0.0)

        # 将范围中点写回标称值（保持兼容：cfg.noise_rms_open 等字段始终有合理值）
        self.noise_rms_open = 0.5 * (self.noise_rms_open_min + self.noise_rms_open_max)
        self.noise_rms_closed = 0.5 * (self.noise_rms_closed_min + self.noise_rms_closed_max)
        self.mech_amp_open = 0.5 * (self.mech_amp_open_min + self.mech_amp_open_max)
        self.mech_amp_closed = 0.5 * (self.mech_amp_closed_min + self.mech_amp_closed_max)
        self.drift_amp_open = 0.5 * (self.drift_amp_open_min + self.drift_amp_open_max)
        self.drift_amp_closed = 0.5 * (self.drift_amp_closed_min + self.drift_amp_closed_max)
        self.drift_freq_hz = 0.5 * (self.drift_freq_hz_min + self.drift_freq_hz_max)
        self.sensor_rms = 0.5 * (self.sensor_rms_min + self.sensor_rms_max)




        def _fix_pair(a, b):
            a = float(a); b = float(b)
            if a > b:
                a, b = b, a
            return a, b

        self.mu_runin_min, self.mu_runin_max = _fix_pair(self.mu_runin_min, self.mu_runin_max)
        self.mu_stable_min, self.mu_stable_max = _fix_pair(self.mu_stable_min, self.mu_stable_max)
        self.mu_severe_min, self.mu_severe_max = _fix_pair(self.mu_severe_min, self.mu_severe_max)

        # 简单裁剪（避免极端输入）
        def _clip_mu(x):
            return max(0.0, min(2.0, float(x)))

        self.mu_runin_min = _clip_mu(self.mu_runin_min)
        self.mu_runin_max = _clip_mu(self.mu_runin_max)
        self.mu_stable_min = _clip_mu(self.mu_stable_min)
        self.mu_stable_max = _clip_mu(self.mu_stable_max)
        self.mu_severe_min = _clip_mu(self.mu_severe_min)
        self.mu_severe_max = _clip_mu(self.mu_severe_max)

        # 确保 severe 上限不小于 stable 上限
        if self.mu_severe_max < self.mu_stable_max:
            self.mu_severe_max = self.mu_stable_max

        # 慢变时间常数：>=0（<=0 表示关闭慢变，保持旧逻辑）
        for _k in ["tau_rpm_s", "tau_mech_s", "tau_noise_s", "tau_sensor_s", "tau_drift_amp_s", "tau_drift_freq_s"]:
            if hasattr(self, _k):
                try:
                    v = float(getattr(self, _k))
                except Exception:
                    v = 0.0
                if not np.isfinite(v):
                    v = 0.0
                setattr(self, _k, max(0.0, v))


    def rpm_range(self) -> Tuple[float, float]:
        """转速范围（rpm）"""
        rmin = float(getattr(self, "rpm_min", self.rpm))
        rmax = float(getattr(self, "rpm_max", self.rpm))
        if rmin > rmax:
            rmin, rmax = rmax, rmin
        return float(rmin), float(rmax)

    def rpm_used(self) -> float:
        """当前用于派生计算的转速（rpm）"""
        rmin, rmax = self.rpm_range()
        return float(0.5 * (rmin + rmax))

    def mech_freq(self) -> float:
        """机械主频（Hz）"""
        return (float(self.rpm_used()) / 60.0) * float(self.mech_harmonic)

    def notch_q_used(self) -> float:
        """陷波 Q（自动）"""
        return _notch_q_from_rpm(self.rpm_used())


def _pick_mu_anchors(cfg: SimConfig, rng: np.random.Generator):
    """
    根据三阶段 μ 范围生成本次仿真的“使用值”
    - 若 min==max => 固定值
    - 使用 rng，保证同一 seed 可重复
    """
    def u(a, b):
        a = float(a); b = float(b)
        if abs(a - b) < 1e-12:
            return float(a)
        lo, hi = (a, b) if a <= b else (b, a)
        return float(rng.uniform(lo, hi))

    mu_st = u(cfg.mu_stable_min, cfg.mu_stable_max)
    mu_r0 = u(cfg.mu_runin_min, cfg.mu_runin_max)
    if mu_r0 < mu_st:
        mu_r0 = min(float(cfg.mu_runin_max), mu_st + 0.05)

    mu_se = u(cfg.mu_severe_min, cfg.mu_severe_max)
    if mu_se < mu_st + 0.05:
        mu_se = min(float(cfg.mu_severe_max), mu_st + 0.10)

    # 最后再裁剪一次
    mu_r0 = float(max(cfg.mu_runin_min, min(cfg.mu_runin_max, mu_r0)))
    mu_st = float(max(cfg.mu_stable_min, min(cfg.mu_stable_max, mu_st)))
    mu_se = float(max(cfg.mu_severe_min, min(cfg.mu_severe_max, mu_se)))
    return mu_r0, mu_st, mu_se


def _mu_profile(t: np.ndarray, cfg: SimConfig, mu_runin_start: float, mu_stable: float, mu_severe_end: float) -> np.ndarray:
    """
    μ(t) 真值：磨合 → 稳定 → 加速磨损（连续）
    说明（重要）：
    - 仅对“锚点”做范围约束（μrunin_start、μstable、μsevere_end），曲线本身用连续映射生成；
      避免对每个采样点做 clip 导致段间边界跳变（不连续）。
    - 三段边界处强制连续：μ(t_runin-) = μ(t_runin+)；μ(t_severe-) = μ(t_severe+)
    """
    T = float(t[-1]) if len(t) else 0.0
    if T <= 0:
        return np.array([], dtype=float)

    # 阶段比例（已在 cfg.validate() 中归一化，这里再做一次保护）
    r1 = max(0.0, float(cfg.phase_runin_ratio))
    r2 = max(0.0, float(cfg.phase_stable_ratio))
    r3 = max(0.0, float(cfg.phase_severe_ratio))
    s = max(1e-9, r1 + r2 + r3)
    r1, r2, r3 = r1 / s, r2 / s, r3 / s

    t_runin = max(0.0, min(T, r1 * T))
    t_severe = max(t_runin, min(T, (r1 + r2) * T))

    # 锚点（仅锚点范围约束，避免段间不连续）
    mu_st0 = float(np.clip(mu_stable, cfg.mu_stable_min, cfg.mu_stable_max))
    mu_r0 = float(np.clip(mu_runin_start, cfg.mu_runin_min, cfg.mu_runin_max))
    mu_se = float(np.clip(mu_severe_end, cfg.mu_severe_min, cfg.mu_severe_max))

    # 约束基本顺序（保证合理单调关系）
    if mu_r0 < mu_st0:
        mu_r0 = mu_st0
    # 稳定段末端：给一个很小的漂移（可在后续加为参数）
    mu_st1 = float(np.clip(mu_st0 + 0.02, cfg.mu_stable_min, cfg.mu_stable_max))
    if mu_se < mu_st1:
        mu_se = mu_st1

    mu = np.empty_like(t, dtype=float)

    # 1) 磨合段：指数衰减（归一化保证 t=t_runin 时刚好到 μst0）
    if t_runin <= 1.0 / max(1e-9, cfg.fs_Hz):
        # 磨合段几乎为0：直接从稳定段开始
        idx1 = t <= t_runin
        if np.any(idx1):
            mu[idx1] = mu_st0
    else:
        idx1 = t <= t_runin
        if np.any(idx1):
            k_rs = float(max(0.05, getattr(cfg, 'trans_runin2stable_k', 1.0)))
            tau = max(1e-6, (0.25 * t_runin) / k_rs)
            e1 = np.exp(-t_runin / tau)
            den = max(1e-12, 1.0 - e1)
            w = (np.exp(-t[idx1] / tau) - e1) / den   # t=0 ->1, t=t_runin ->0
            mu[idx1] = mu_st0 + (mu_r0 - mu_st0) * w

    # 2) 稳定段：轻微漂移（线性，保证边界连续）
    idx2 = (t > t_runin) & (t <= t_severe)
    if np.any(idx2):
        span = max(1e-6, (t_severe - t_runin))
        x = (t[idx2] - t_runin) / span  # 0..1
        mu[idx2] = mu_st0 + (mu_st1 - mu_st0) * x

    # 3) 加速磨损：S型上升（归一化 sigmoid，保证起点=μst1，终点=μse）
    idx3 = t > t_severe
    if np.any(idx3):
        span = max(1e-6, (T - t_severe))
        x = (t[idx3] - t_severe) / span  # 0..1
        k_sa = float(max(0.05, getattr(cfg, 'trans_stable2severe_k', 1.0)))
        k = 10.0 * k_sa
        sig = 1.0 / (1.0 + np.exp(-k * (x - 0.5)))
        s0 = 1.0 / (1.0 + np.exp(-k * (0.0 - 0.5)))
        s1 = 1.0 / (1.0 + np.exp(-k * (1.0 - 0.5)))
        den = max(1e-12, (s1 - s0))
        sn = (sig - s0) / den  # x=0 ->0, x=1 ->1
        mu[idx3] = mu_st1 + (mu_se - mu_st1) * sn

    # 微弱高频成分（非机械主频）
    mu += 0.002 * np.sin(2 * np.pi * 0.2 * t)
    return mu


def _rand_in_range(a: float, b: float, rng: np.random.Generator) -> float:
    a = float(a); b = float(b)
    if a > b:
        a, b = b, a
    if abs(a - b) < 1e-15:
        return a
    return float(rng.uniform(a, b))


def _slow_bounded_series(
    t: np.ndarray,
    vmin: float,
    vmax: float,
    tau_s: float,
    rng: np.random.Generator,
    sigma_z: float = 0.35,
    z0: float | None = None,
) -> np.ndarray:
    """生成在 [vmin, vmax] 内缓慢变化的随机序列。

    机制：
    - 在无界变量 z 上做 OU 慢漂移（时间常数 tau_s）
    - 再用 tanh 压缩映射到 [vmin, vmax]，保证永不越界
    """
    vmin = float(vmin); vmax = float(vmax)
    if vmin > vmax:
        vmin, vmax = vmax, vmin
    if len(t) == 0:
        return np.array([], dtype=float)

    # 退化情形：不启用慢变 / 范围退化 -> 常量
    if tau_s is None:
        tau_s = 0.0
    tau_s = float(tau_s)
    if tau_s <= 0.0 or abs(vmax - vmin) < 1e-15:
        v0 = _rand_in_range(vmin, vmax, rng)
        return np.full_like(t, v0, dtype=float)

    # 以均匀 dt 近似（仿真本身就是等间隔采样）
    if len(t) >= 2:
        dt = float(np.median(np.diff(t)))
    else:
        dt = 1.0

    dt = max(dt, 1e-6)
    # OU: z_{k+1} = a z_k + b N(0,1)
    a = math.exp(-dt / tau_s)
    b = math.sqrt(max(0.0, 1.0 - a * a)) * sigma_z

    if z0 is None:
        z = float(rng.normal(0.0, 0.8))
    else:
        z = float(z0)

    out = np.empty_like(t, dtype=float)
    span = vmax - vmin
    for i in range(len(t)):
        # 映射到 [vmin, vmax]
        u = 0.5 * (math.tanh(z) + 1.0)  # in (0,1)
        out[i] = vmin + span * u
        # 演化
        z = a * z + b * float(rng.normal(0.0, 1.0))
    return out


def _make_quasi_drift(
    t: np.ndarray,
    amp_min: float,
    amp_max: float,
    f_min: float,
    f_max: float,
    rng: np.random.Generator,
    tau_amp_s: float = 0.0,
    tau_freq_s: float = 0.0,
) -> np.ndarray:
    """生成低频准周期漂移：两频率分量叠加（beat/准周期）。

    相比旧版“整段只抽一次 A/f”，这里支持 A(t)、f(t) 在范围内慢变：
    - tau_amp_s / tau_freq_s <= 0 时退化为旧逻辑（兼容）
    """
    # 幅值包络 A(t)
    A_t = _slow_bounded_series(t, amp_min, amp_max, tau_amp_s, rng)

    # 两个频率分量：f1(t), f2(t)（都在范围内慢变）
    f1_t = _slow_bounded_series(t, f_min, f_max, tau_freq_s, rng, sigma_z=0.25, z0=float(rng.normal(0.0, 0.7)))
    f2_t = _slow_bounded_series(t, f_min, f_max, tau_freq_s, rng, sigma_z=0.25, z0=float(rng.normal(0.0, 0.7)))

    # 防止两者过于接近导致“看起来像单频”
    span_f = float(max(1e-12, abs(f_max - f_min)))
    sep = 0.05 * span_f
    diff = np.abs(f2_t - f1_t)
    if np.any(diff < sep):
        # 对过近的点做轻微偏移并裁剪回范围
        sign = np.where((f2_t - f1_t) >= 0.0, 1.0, -1.0)
        f2_t = np.clip(f2_t + sign * sep, min(f_min, f_max), max(f_min, f_max))

    # 幅值分配（随时间慢变的 A_t 作为总包络）
    w = float(rng.uniform(0.55, 0.80))
    a1_t = A_t * w
    a2_t = np.maximum(0.0, A_t - a1_t) * float(rng.uniform(0.70, 1.20))

    # 相位：对瞬时频率积分
    if len(t) >= 2:
        dt = float(np.median(np.diff(t)))
    else:
        dt = 1.0
    dt = max(dt, 1e-6)

    p1 = float(rng.uniform(0.0, 2.0 * np.pi))
    p2 = float(rng.uniform(0.0, 2.0 * np.pi))
    phi1 = 2.0 * np.pi * np.cumsum(f1_t) * dt + p1
    phi2 = 2.0 * np.pi * np.cumsum(f2_t) * dt + p2

    return a1_t * np.sin(phi1) + a2_t * np.sin(phi2)



def _make_tavg_open_closed(t: np.ndarray, cfg: SimConfig, rng: np.random.Generator):
    """
    开环/闭环平均张力 T_avg(t)
    说明：不做逐步 PID 微步仿真，而是构造“开环=扰动明显，闭环=扰动衰减”的等效效果。

    2026-02 更新：
    - 机械周期扰动幅值 A_mech(t)、高频噪声 RMS(t)、低频漂移 A_drift(t)/f_drift(t)
      支持在给定范围内“随时间慢变”（更贴近真实准周期/工况漂移）。
    - 慢变速度由 cfg.tau_*_s 控制；tau<=0 时自动退化为旧逻辑（整段抽一次常量）。
    """
    # 转速慢变：rpm(t) -> f_mech(t)
    rpm_min, rpm_max = cfg.rpm_range()
    rpm_t = _slow_bounded_series(t, rpm_min, rpm_max, getattr(cfg, "tau_rpm_s", 0.0), rng, sigma_z=0.25)
    f_mech_t = (rpm_t / 60.0) * float(max(1, cfg.mech_harmonic))

    # 机械周期扰动幅值（慢变）
    mech_amp_open_t = _slow_bounded_series(
        t, cfg.mech_amp_open_min, cfg.mech_amp_open_max, getattr(cfg, "tau_mech_s", 0.0), rng
    )
    mech_amp_closed_t = _slow_bounded_series(
        t, cfg.mech_amp_closed_min, cfg.mech_amp_closed_max, getattr(cfg, "tau_mech_s", 0.0), rng
    )

    # 高频噪声 RMS（慢变）
    noise_rms_open_t = _slow_bounded_series(
        t, cfg.noise_rms_open_min, cfg.noise_rms_open_max, getattr(cfg, "tau_noise_s", 0.0), rng
    )
    noise_rms_closed_t = _slow_bounded_series(
        t, cfg.noise_rms_closed_min, cfg.noise_rms_closed_max, getattr(cfg, "tau_noise_s", 0.0), rng
    )

    # 低频准周期漂移（beat）：幅值/频率均可慢变
    drift_open = _make_quasi_drift(
        t, cfg.drift_amp_open_min, cfg.drift_amp_open_max,
        cfg.drift_freq_hz_min, cfg.drift_freq_hz_max, rng,
        tau_amp_s=getattr(cfg, "tau_drift_amp_s", 0.0),
        tau_freq_s=getattr(cfg, "tau_drift_freq_s", 0.0),
    )
    drift_closed = _make_quasi_drift(
        t, cfg.drift_amp_closed_min, cfg.drift_amp_closed_max,
        cfg.drift_freq_hz_min, cfg.drift_freq_hz_max, rng,
        tau_amp_s=getattr(cfg, "tau_drift_amp_s", 0.0),
        tau_freq_s=getattr(cfg, "tau_drift_freq_s", 0.0),
    )

    # 频率随时间变化时，用相位累积生成准周期信号
    if len(t) >= 2:
        dt = float(np.median(np.diff(t)))
    else:
        dt = 1.0
    dt = max(dt, 1e-6)
    p1 = float(rng.uniform(0.0, 2.0 * np.pi))
    p2 = float(rng.uniform(0.0, 2.0 * np.pi))
    phi1 = 2.0 * np.pi * np.cumsum(f_mech_t) * dt + p1
    phi2 = 2.0 * np.pi * np.cumsum(2.0 * f_mech_t) * dt + p2

    mech_open = mech_amp_open_t * np.sin(phi1) \
                + 0.3 * mech_amp_open_t * np.sin(phi2 + 0.9)
    mech_closed = mech_amp_closed_t * np.sin(phi1) \
                  + 0.3 * mech_amp_closed_t * np.sin(phi2 + 0.9)

    noise_open = rng.normal(0.0, 1.0, size=len(t)) * noise_rms_open_t
    noise_closed = rng.normal(0.0, 1.0, size=len(t)) * noise_rms_closed_t

    t_open = cfg.t_set_N + drift_open + mech_open + noise_open

    phi_res = 2.0 * np.pi * np.cumsum(0.5 * f_mech_t) * dt + 1.1
    residual = 0.02 * cfg.t_set_N * np.sin(phi_res)
    t_closed = cfg.t_set_N + drift_closed + mech_closed + noise_closed + residual

    return t_open, t_closed, rpm_t, f_mech_t


def _tensions_from_tavg_mu(t: np.ndarray, tavg: np.ndarray, mu: np.ndarray, cfg: SimConfig, rng: np.random.Generator):
    """
    由 T_avg 与 μ 生成紧边/松边张力：
    r = exp(μθ)，且 (T_high+T_low)/2 = T_avg
    => T_high = 2*T_avg*r/(1+r),  T_low = 2*T_avg/(1+r)
    """
    theta = np.deg2rad(cfg.theta_deg)
    r = np.exp(np.clip(mu * theta, -10, 10))
    t_high = 2.0 * tavg * r / (1.0 + r)
    t_low = 2.0 * tavg / (1.0 + r)

    sensor_rms_t = _slow_bounded_series(t, cfg.sensor_rms_min, cfg.sensor_rms_max, getattr(cfg, "tau_sensor_s", 0.0), rng)

    t_high = np.maximum(t_high + rng.normal(0.0, 1.0, size=len(t_high)) * sensor_rms_t, 0.0)
    t_low = np.maximum(t_low + rng.normal(0.0, 1.0, size=len(t_low)) * sensor_rms_t, 0.0)
    return t_high, t_low


def _invert_mu_from_tensions(t_high: np.ndarray, t_low: np.ndarray, cfg: SimConfig):
    """Capstan 反演：μ = ln(T_high/T_low)/θ，并进行门控/裁剪"""
    theta = np.deg2rad(cfg.theta_deg)
    eps = 1e-9
    q = (t_high >= cfg.tmin_gate_N) & (t_low >= cfg.tmin_gate_N)
    ratio = (t_high + eps) / (t_low + eps)
    ratio = np.clip(ratio, cfg.ratio_clip_min, cfg.ratio_clip_max)
    mu = np.log(ratio) / max(theta, 1e-9)
    mu = mu.astype(float)
    mu[~q] = np.nan
    return mu, q.astype(int)


def _find_stable_baseline(mu_f: np.ndarray, t: np.ndarray, q_valid: np.ndarray, cfg: SimConfig, end_idx: int = None):
    """
    稳定段基线 μss（按你的新规则）：
    - 稳定窗口长度 W_ss -> N_ss 不变
    - 窗口有效性条件不变（有效比例 + std + slope）
    - 窗口滑动扫描逻辑不变（step=win//10）
    - 先形成连续稳定窗口并集（stable segments），再用最短连续稳定段 Whold 过滤
    - μss = 第一个“连续稳定段”（满足 Whold）的中位数（median），而非第一个稳定窗口
    返回：((k_start, k_end), mu_ss)
    end_idx：仅在 [0,end_idx) 范围内判定（用于“超限后不再判定稳定段”）
    """
    segs = _find_stable_segments(mu_f, t, q_valid, cfg, end_idx=end_idx)
    if not segs:
        return None, None
    k0, k1 = segs[0]
    yy = mu_f[k0:k1]
    qq = q_valid[k0:k1]
    m = np.isfinite(yy) & (qq.astype(bool))
    if int(m.sum()) < 5:
        m = np.isfinite(yy)
    mu_ss = float(np.nanmedian(yy[m])) if int(np.isfinite(yy[m]).sum()) > 0 else None
    return (k0, k1), mu_ss
def _find_stable_segments(mu_f: np.ndarray, t: np.ndarray, q_valid: np.ndarray, cfg: SimConfig, end_idx: int = None) -> List[Tuple[int, int]]:
    """
    连续稳定段窗口识别（用于可视化）：
    - 用与 _find_stable_baseline 相同的窗口/判据扫描
    - 将满足条件的窗口并集形成稳定段 mask
    - 返回合并后的连续区间列表：[(k_start, k_end), ...] 其中 k_end 为开区间
    - end_idx：仅在 [0,end_idx) 范围内判定（超限后不再判定）
    """
    n_all = len(mu_f)
    if n_all == 0:
        return []
    n = n_all if (end_idx is None) else int(max(0, min(n_all, end_idx)))
    if n <= 0:
        return []

    win = int(round(cfg.stable_win_s * cfg.fs_Hz))
    win = max(win, 50)
    step = max(1, win // 10)

    stable_mask = np.zeros(n_all, dtype=bool)

    for k0 in range(0, n - win, step):
        k1 = k0 + win
        if q_valid[k0:k1].mean() < cfg.stable_valid_min:
            continue
        seg = mu_f[k0:k1]
        seg = seg[np.isfinite(seg)]
        if len(seg) < 10:
            continue
        if float(np.std(seg)) > cfg.stable_sigma_max:
            continue

        # 总漂移量判据（替代斜率判据）：
        # 取窗口前/后 10% 样本的中位数差值作为漂移量 Δμ，并与阈值比较（Δμ_max=stable_slope_max）
        yy = mu_f[k0:k1]
        vv = q_valid[k0:k1].astype(bool)
        mm = np.isfinite(yy) & vv
        if mm.sum() < 10:
            mm = np.isfinite(yy)
        if mm.sum() < 10:
            continue
        vals = yy[mm]
        nvals = int(len(vals))
        edge = max(5, int(round(0.10 * nvals)))
        edge = min(edge, max(5, nvals // 2))
        head = np.nanmedian(vals[:edge])
        tail = np.nanmedian(vals[-edge:])
        drift = float(abs(tail - head))
        if drift > cfg.stable_slope_max:
            continue

        stable_mask[k0:k1] = True

    segs: List[Tuple[int, int]] = []
    i = 0
    while i < n:
        if stable_mask[i]:
            j = i + 1
            while j < n and stable_mask[j]:
                j += 1
            segs.append((i, j))
            i = j
        else:
            i += 1
    # 依据 Whold 过滤：连续稳定窗口并集覆盖时间不足者不计为稳定段
    min_hold_s = float(getattr(cfg, 'stable_hold_s', 0.0))
    if min_hold_s > 0.0:
        segs2: List[Tuple[int, int]] = []
        for a, b in segs:
            dur_s = float(max(0.0, (b - a) / max(1e-9, cfg.fs_Hz)))
            if dur_s + 1e-9 >= min_hold_s:
                segs2.append((a, b))
        segs = segs2
    return segs


def _find_failure_time(mu_f: np.ndarray, t: np.ndarray, mu_ss: Optional[float], cfg: SimConfig, start_idx: int = 0):
    """寿命判据：μ 持续超过 μth=μss*(1+δ) 持续 Wpersist，输出 tlife"""
    if mu_ss is None or not np.isfinite(mu_ss):
        return None, None
    mu_th = float(mu_ss * (1.0 + cfg.fail_delta))
    hold = int(round(cfg.fail_hold_s * cfg.fs_Hz))
    hold = max(1, hold)

    # 关键修复：失效判定从 start_idx 开始（通常取稳定段基线窗口结束），避免磨合段 μ 偏高导致 tlife=0s
    start_idx = int(max(0, min(len(mu_f) - 1, start_idx)))

    above = np.isfinite(mu_f) & (mu_f > mu_th)
    count = 0
    for i in range(start_idx, len(above)):
        a = bool(above[i])
        if a:
            count += 1
            if count >= hold:
                tlife = float(t[i - hold + 1])
                return tlife, mu_th
        else:
            count = 0
    return None, mu_th


def simulate(cfg: SimConfig, seed: int = 7, progress_cb: ProgressCB = None) -> Dict[str, Any]:
    """仅生成仿真结果（不导出文件）"""
    cfg.validate()
    rng = np.random.default_rng(seed)

    _cb(progress_cb, 5.0, "生成时间轴...")
    n = int(round(cfg.fs_Hz * cfg.duration_s))
    n = max(n, 2)
    t = np.arange(n, dtype=float) / float(cfg.fs_Hz)

    _cb(progress_cb, 12.0, "生成 μ(t) 真值...")
    mu_runin_start_used, mu_stable_used, mu_severe_end_used = _pick_mu_anchors(cfg, rng)
    mu_true = _mu_profile(t, cfg, mu_runin_start_used, mu_stable_used, mu_severe_end_used)

    _cb(progress_cb, 18.0, "生成开环/闭环平均张力...")
    tavg_open, tavg_closed, rpm_t, f_mech_t = _make_tavg_open_closed(t, cfg, rng)

    _cb(progress_cb, 24.0, "生成高/低张力侧张力（含传感器噪声）...")
    th_open, tl_open = _tensions_from_tavg_mu(t, tavg_open, mu_true, cfg, rng)
    th_cl, tl_cl = _tensions_from_tavg_mu(t, tavg_closed, mu_true, cfg, rng)

    ff_open = th_open - tl_open
    ff_cl = th_cl - tl_cl

    _cb(progress_cb, 30.0, "Capstan 反演 μ_raw...")
    mu_open_raw, q_open = _invert_mu_from_tensions(th_open, tl_open, cfg)
    mu_cl_raw, q_cl = _invert_mu_from_tensions(th_cl, tl_cl, cfg)

    _cb(progress_cb, 36.0, "μ 异常点剔除（Hampel）...")
    hampel_win = int(round(cfg.hampel_win_s * cfg.fs_Hz))
    hampel_win = max(1, hampel_win)
    mu_cl_h = _hampel_filter_nan(mu_cl_raw, hampel_win, cfg.hampel_nsig)

    fill = np.nanmedian(mu_cl_h[np.isfinite(mu_cl_h)]) if np.any(np.isfinite(mu_cl_h)) else 0.0
    mu_f0 = np.nan_to_num(mu_cl_h, nan=fill)

    rpm_used = float(np.nanmean(rpm_t)) if len(rpm_t) > 0 else cfg.rpm_used()
    f_mech_used = (rpm_used / 60.0) * float(max(1, cfg.mech_harmonic))
    q_notch = _notch_q_from_rpm(rpm_used)
    q_notch_t = _notch_q_from_rpm_vec(rpm_t) if len(rpm_t) > 0 else np.full_like(mu_f0, q_notch, dtype=float)

    rpm_span = float(np.nanmax(rpm_t) - np.nanmin(rpm_t)) if len(rpm_t) > 0 else 0.0
    use_tv_notch = (rpm_span > 1e-6) and (float(getattr(cfg, "tau_rpm_s", 0.0)) > 0.0)

    if use_tv_notch:
        f_min = float(np.nanmin(f_mech_t))
        f_max = float(np.nanmax(f_mech_t))
        q_min = float(np.nanmin(q_notch_t))
        q_max = float(np.nanmax(q_notch_t))
        _cb(progress_cb, 40.0, f"μ 陷波（f_mech∈[{f_min:.6g},{f_max:.6g}]Hz, Q∈[{q_min:.3g},{q_max:.3g}]）...")
        block_s = min(30.0, max(2.0, float(getattr(cfg, "tau_rpm_s", 0.0)) / 10.0))
        mu_cl_notch = _iir_notch(mu_f0, fs=cfg.fs_Hz, f0=f_mech_t, q=q_notch_t, block_s=block_s)
    else:
        _cb(progress_cb, 40.0, f"μ 陷波（f_mech={f_mech_used:.6g}Hz, Q={q_notch:.3g}）...")
        mu_cl_notch = _iir_notch(mu_f0, fs=cfg.fs_Hz, f0=f_mech_used, q=q_notch)

    _cb(progress_cb, 44.0, "μ 低通滤波...")
    mu_cl_lp = _lowpass(mu_cl_notch, fs=cfg.fs_Hz, fc=cfg.lowpass_fc_hz, order=3).astype(float)
    mu_cl_lp[q_cl == 0] = np.nan

    _cb(progress_cb, 48.0, "稳定段基线与寿命判据计算...")
    # 1) 先识别“全时域”的连续稳定段（用于可视化）；绘图时再裁剪 tlife 右侧
    stable_segs_all = _find_stable_segments(mu_cl_lp, t, q_cl, cfg, end_idx=None)

    # 2) 先在全时域寻找稳定段基线 μss（得到稳定窗口 stable_idx0）
    stable_idx0, mu_ss0 = _find_stable_baseline(mu_cl_lp, t, q_cl, cfg, end_idx=None)

    # 3) 初步计算 tlife（从稳定窗口结束开始判定，避免磨合段 μ 偏高导致 tlife=0s）
    start_fail_idx0 = int(stable_idx0[1]) if (stable_idx0 is not None) else 0
    tlife0, mu_th0 = _find_failure_time(mu_cl_lp, t, mu_ss0, cfg, start_idx=start_fail_idx0)

    # 4) 若存在 tlife0，则优先在 tlife0 之前重算 μss（更符合“失效前基线”意义）；否则沿用全时域
    tlife = tlife0
    mu_th = mu_th0
    tlife_idx = None
    if tlife0 is not None:
        tlife_idx = int(max(1, min(len(t), int(round(float(tlife0) * cfg.fs_Hz)) + 1)))
        stable_idx, mu_ss = _find_stable_baseline(mu_cl_lp, t, q_cl, cfg, end_idx=tlife_idx)
        if mu_ss is None or not np.isfinite(mu_ss):
            stable_idx, mu_ss = stable_idx0, mu_ss0
    else:
        stable_idx, mu_ss = stable_idx0, mu_ss0

    # 5) 用最终 μss 重新计算 tlife 与 μth（同样从稳定窗口结束开始）
    if mu_ss is not None and np.isfinite(mu_ss):
        start_fail_idx = int(stable_idx[1]) if (stable_idx is not None) else 0
        tlife, mu_th = _find_failure_time(mu_cl_lp, t, mu_ss, cfg, start_idx=start_fail_idx)
        if tlife is not None:
            tlife_idx = int(max(1, min(len(t), int(round(float(tlife) * cfg.fs_Hz)) + 1)))

    stable_segs = stable_segs_all

    res = {
        "cfg": cfg,
        "seed": int(seed),
        "derived": {
            "rpm_min": float(cfg.rpm_min),
            "rpm_max": float(cfg.rpm_max),
            "rpm_used": float(rpm_used),
            "mech_freq_hz_used": float(f_mech_used),
            "notch_q_used": float(q_notch),
            "mu_runin_start_used": float(mu_runin_start_used),
            "mu_stable_used": float(mu_stable_used),
            "mu_severe_end_used": float(mu_severe_end_used),
            "phase_runin_ratio_used": float(cfg.phase_runin_ratio),
            "phase_stable_ratio_used": float(cfg.phase_stable_ratio),
            "phase_severe_ratio_used": float(cfg.phase_severe_ratio),
        },
        "series": {
            "t": t,
            "mu_true": mu_true,
            "rpm_t": rpm_t,
            "f_mech_t": f_mech_t,
            "q_notch_t": q_notch_t,
            "open": {"tavg": tavg_open, "th": th_open, "tl": tl_open, "ff": ff_open, "mu_raw": mu_open_raw, "q": q_open},
            "closed": {
                "tavg": tavg_closed, "th": th_cl, "tl": tl_cl, "ff": ff_cl,
                "mu_raw": mu_cl_raw, "mu_hampel": mu_cl_h, "mu_notch": mu_cl_notch, "mu_filt": mu_cl_lp, "q": q_cl
            },
        },
        "baseline": {"stable_window_idx": stable_idx, "stable_segments_idx": stable_segs, "mu_ss": mu_ss, "tlife_idx": tlife_idx},
        "threshold": {"mu_th": mu_th},
        "life": {"tlife_s": tlife},
    }
    _cb(progress_cb, 50.0, "仿真结果生成完成（可导出 xlsx 或导出图片）")
    return res


def export_xlsx(res: Dict[str, Any], out_dir: str, progress_cb: ProgressCB = None) -> str:
    """导出 xlsx（多sheet自动拆分）"""
    cfg: SimConfig = res["cfg"]
    out_dir = os.path.abspath(out_dir)
    _ensure_dir(out_dir)

    _cb(progress_cb, 52.0, "组织导出数据表...")
    t = res["series"]["t"]
    stride = int(cfg.export_stride)
    t_e = t[::stride]

    def ds(x): return x[::stride]
    f_mech = float(res["derived"]["mech_freq_hz_used"])
    q_notch = float(res["derived"]["notch_q_used"])
    rpm_t = res["series"].get("rpm_t")
    f_mech_t = res["series"].get("f_mech_t")
    q_notch_t = res["series"].get("q_notch_t")
    if rpm_t is None:
        rpm_t = np.full_like(t, float(cfg.rpm_used()), dtype=float)
    if f_mech_t is None:
        f_mech_t = (rpm_t / 60.0) * float(max(1, cfg.mech_harmonic))
    if q_notch_t is None:
        q_notch_t = np.full_like(t, q_notch, dtype=float)

    cl = res["series"]["closed"]
    op = res["series"]["open"]

    closed_df = pd.DataFrame({
        "t_s": t_e,
        "t_high_N": ds(cl["th"]),
        "t_low_N": ds(cl["tl"]),
        "t_avg_N": ds(cl["tavg"]),
        "f_fric_N": ds(cl["ff"]),
        "mu_true": ds(res["series"]["mu_true"]),
        "mu_runin_start_used": np.full_like(t_e, float(res["derived"].get("mu_runin_start_used", np.nan)), dtype=float),
        "mu_stable_used": np.full_like(t_e, float(res["derived"].get("mu_stable_used", np.nan)), dtype=float),
        "mu_severe_end_used": np.full_like(t_e, float(res["derived"].get("mu_severe_end_used", np.nan)), dtype=float),
        "phase_runin_ratio": np.full_like(t_e, float(res["derived"].get("phase_runin_ratio_used", np.nan)), dtype=float),
        "phase_stable_ratio": np.full_like(t_e, float(res["derived"].get("phase_stable_ratio_used", np.nan)), dtype=float),
        "phase_severe_ratio": np.full_like(t_e, float(res["derived"].get("phase_severe_ratio_used", np.nan)), dtype=float),
        "mu_raw": ds(cl["mu_raw"]),
        "mu_hampel": ds(cl["mu_hampel"]),
        "mu_notch": ds(cl["mu_notch"]),
        "mu_filt": ds(cl["mu_filt"]),
        "q_valid": ds(cl["q"]).astype(int),
        "rpm": ds(rpm_t),
        "mech_freq_hz": ds(f_mech_t),
        "notch_q_used": ds(q_notch_t),
        "theta_deg": np.full_like(t_e, float(cfg.theta_deg), dtype=float),
        "t_set_N": np.full_like(t_e, float(cfg.t_set_N), dtype=float),
    })

    open_df = pd.DataFrame({
        "t_s": t_e,
        "t_high_N": ds(op["th"]),
        "t_low_N": ds(op["tl"]),
        "t_avg_N": ds(op["tavg"]),
        "f_fric_N": ds(op["ff"]),
        "mu_true": ds(res["series"]["mu_true"]),
        "mu_runin_start_used": np.full_like(t_e, float(res["derived"].get("mu_runin_start_used", np.nan)), dtype=float),
        "mu_stable_used": np.full_like(t_e, float(res["derived"].get("mu_stable_used", np.nan)), dtype=float),
        "mu_severe_end_used": np.full_like(t_e, float(res["derived"].get("mu_severe_end_used", np.nan)), dtype=float),
        "phase_runin_ratio": np.full_like(t_e, float(res["derived"].get("phase_runin_ratio_used", np.nan)), dtype=float),
        "phase_stable_ratio": np.full_like(t_e, float(res["derived"].get("phase_stable_ratio_used", np.nan)), dtype=float),
        "phase_severe_ratio": np.full_like(t_e, float(res["derived"].get("phase_severe_ratio_used", np.nan)), dtype=float),
        "mu_raw": ds(op["mu_raw"]),
        "q_valid": ds(op["q"]).astype(int),
        "rpm": ds(rpm_t),
        "mech_freq_hz": ds(f_mech_t),
        "notch_q_used": ds(q_notch_t),
        "theta_deg": np.full_like(t_e, float(cfg.theta_deg), dtype=float),
        "t_set_N": np.full_like(t_e, float(cfg.t_set_N), dtype=float),
    })

    
    # 开环-闭环对比输出范围数据表（可选）
    t0_cmp = float(getattr(cfg, "compare_t_start_s", 0.0))
    t1_cmp = float(getattr(cfg, "compare_t_end_s", -1.0))
    if (t0_cmp == 0.0) and (t1_cmp < 0.0):
        cmp_mask = np.ones_like(t_e, dtype=bool)
    else:
        if t1_cmp < 0.0:
            t1_cmp = float(t_e[-1])
        t0_cmp = max(0.0, t0_cmp)
        t1_cmp = max(t0_cmp, t1_cmp)
        cmp_mask = (t_e >= t0_cmp) & (t_e <= t1_cmp)
        if int(cmp_mask.sum()) < 2:
            cmp_mask = np.ones_like(t_e, dtype=bool)

    compare_df = pd.DataFrame({
        "t_s": t_e[cmp_mask],
        "t_avg_open_N": ds(op["tavg"])[cmp_mask],
        "t_avg_closed_N": ds(cl["tavg"])[cmp_mask],
    })

    # 列名添加中文括号注释
    col_note = {
        "t_s": "t_s（时间s）",
        "t_high_N": "t_high_N（高张力N）",
        "t_low_N": "t_low_N（低张力N）",
        "t_avg_N": "t_avg_N（平均张力N）",
        "f_fric_N": "f_fric_N（摩擦力N）",
        "mu_true": "mu_true（μ真值）",
        "mu_runin_start_used": "mu_runin_start_used（磨合μ起始）",
        "mu_stable_used": "mu_stable_used（稳定μ）",
        "mu_severe_end_used": "mu_severe_end_used（加速μ末值）",
        "phase_runin_ratio": "phase_runin_ratio（磨合比例）",
        "phase_stable_ratio": "phase_stable_ratio（稳定比例）",
        "phase_severe_ratio": "phase_severe_ratio（加速比例）",
        "mu_raw": "mu_raw（μ原始）",
        "mu_hampel": "mu_hampel（μ去毛刺）",
        "mu_notch": "mu_notch（μ陷波）",
        "mu_filt": "mu_filt（μ低通）",
        "q_valid": "q_valid（有效标记）",
        "rpm": "rpm（转速rpm）",
        "mech_freq_hz": "mech_freq_hz（主频Hz）",
        "notch_q_used": "notch_q_used（陷波Q）",
        "theta_deg": "theta_deg（包角deg）",
        "t_set_N": "t_set_N（设定张力N）",
        "t_avg_open_N": "t_avg_open_N（开环平均张力N）",
        "t_avg_closed_N": "t_avg_closed_N（闭环平均张力N）",
    }

    closed_df.rename(columns=col_note, inplace=True)
    open_df.rename(columns=col_note, inplace=True)
    compare_df.rename(columns=col_note, inplace=True)

    xlsx_path = os.path.join(out_dir, "needle_hook_wear_sim.xlsx")
    _cb(progress_cb, 55.0, "开始写入 xlsx（多sheet，必要时自动拆分）...")
    _write_xlsx_multisheet(
        xlsx_path,
        {"closed_loop": closed_df, "open_loop": open_df, "compare_window": compare_df},
        progress_cb=progress_cb,
        base_pct=55.0,
        span_pct=44.0
    )
    _cb(progress_cb, 100.0, "xlsx 导出完成")
    return xlsx_path


def _legend_below_center(ncol: int = 3):
    """图例放到下方居中（不在图框内）"""
    plt.legend(loc="upper center", bbox_to_anchor=(0.5, -0.18), ncol=ncol, frameon=False)


def export_plots(
    res: Dict[str, Any],
    out_dir: str,
    lang: str = "zh",
    progress_cb: ProgressCB = None,
    font_prepared: bool = False,
) -> Dict[str, str]:
    """
    导出图片：
    - lang: "zh" 或 "en"
    - 图例：放在下方居中（不在图框内）
    - 稳定段：显示“连续稳定段窗口”（并集），且仅来自“第一次超限前”
    - tlife 数值：只在图例显示（不在图中显示）
    """
    cfg: SimConfig = res["cfg"]
    out_dir = os.path.abspath(out_dir)
    plot_dir = os.path.join(out_dir, "plots")
    _ensure_dir(plot_dir)

    if not font_prepared:
        setup_plot_font(lang=lang, progress_cb=progress_cb, pct=52.0)

    t = res["series"]["t"]
    cl = res["series"]["closed"]
    op = res["series"]["open"]

    mu_f = cl["mu_filt"]
    mu_ss = res["baseline"].get("mu_ss")
    stable_segs = res["baseline"].get("stable_segments_idx") or []
    mu_th = res["threshold"].get("mu_th")
    tlife = res["life"].get("tlife_s")

    # 依据 tlife 裁剪稳定段：先画全时域稳定段，再删除 tlife 右侧稳定段
    stable_segs_clip = []
    if tlife is None:
        stable_segs_clip = list(stable_segs)
    else:
        t_cut = float(tlife)
        for k0, k1 in stable_segs:
            t0 = float(t[k0])
            t1 = float(t[k1]) if (k1 < len(t)) else float(t[-1])
            if t0 >= t_cut:
                continue
            if t1 > t_cut:
                # 裁剪到 tlife
                k1c = int(min(len(t)-1, max(k0+1, round(t_cut * cfg.fs_Hz))))
                stable_segs_clip.append((k0, k1c))
            else:
                stable_segs_clip.append((k0, k1))

    if lang.lower().startswith("en"):
        L = {
            "t": "Time t (s)",
            "tension": "Tension (N)",
            "tavg": "Mean tension (N)",
            "mu": "Friction coefficient μ",
            "title_closed_t": "Closed-loop: Tension vs Time",
            "title_open_t": "Open-loop: Tension vs Time",
            "title_mu": "Closed-loop: μ vs Time (μss & μth)",
            "title_cmp": "Open-loop vs Closed-loop: Mean Tension",
            "th": "T_high",
            "tl": "T_low",
            "ta": "T_avg",
            "mu_f": "μ (filtered)",
            "mu_ss": "Baseline μss",
            "mu_th": "Threshold μth",
            "ss": "Stable segments",
            "tlife": "tlife",
            "open": "Open-loop: T_avg",
            "closed": "Closed-loop: T_avg",
        }
    else:
        L = {
            "t": "时间 t (s)",
            "tension": "张力 (N)",
            "tavg": "平均张力 (N)",
            "mu": "摩擦系数 μ",
            "title_closed_t": "闭环：张力-时间",
            "title_open_t": "开环：张力-时间",
            "title_mu": "闭环：摩擦系数-时间（μss 与 μth）",
            "title_cmp": "开环 vs 闭环：平均张力稳定性对比",
            "th": "高张力侧 T_high",
            "tl": "低张力侧 T_low",
            "ta": "平均张力 T_avg",
            "mu_f": "μ (滤波后)",
            "mu_ss": "稳定段基线 μss",
            "mu_th": "超限阈值 μth",
            "ss": "连续稳定段",
            "tlife": "tlife",
            "open": "开环：T_avg",
            "closed": "闭环：T_avg",
        }

    # 1) 闭环张力-时间
    _cb(progress_cb, 60.0, "绘图：闭环张力-时间...")
    tx, y1 = _downsample_for_plot(t, cl["th"], cfg.plot_max_points)
    _, y2 = _downsample_for_plot(t, cl["tl"], cfg.plot_max_points)
    _, y3 = _downsample_for_plot(t, cl["tavg"], cfg.plot_max_points)
    plt.figure()
    plt.plot(tx, y1, label=L["th"], color="tab:blue")
    plt.plot(tx, y2, label=L["tl"], color="tab:orange")
    plt.plot(tx, y3, label=L["ta"], color="tab:green")
    plt.xlabel(L["t"])
    plt.ylabel(L["tension"])
    plt.title(L["title_closed_t"])
    _legend_below_center(ncol=3)
    plt.tight_layout(rect=[0, 0.10, 1, 1])
    p_closed = os.path.join(plot_dir, f"closed_tensions_{lang}.png")
    plt.savefig(p_closed, dpi=180)
    plt.close()

    # 2) 开环张力-时间
    _cb(progress_cb, 67.0, "绘图：开环张力-时间...")
    tx, y1 = _downsample_for_plot(t, op["th"], cfg.plot_max_points)
    _, y2 = _downsample_for_plot(t, op["tl"], cfg.plot_max_points)
    _, y3 = _downsample_for_plot(t, op["tavg"], cfg.plot_max_points)
    plt.figure()
    plt.plot(tx, y1, label=L["th"], color="tab:blue")
    plt.plot(tx, y2, label=L["tl"], color="tab:orange")
    plt.plot(tx, y3, label=L["ta"], color="tab:green")
    plt.xlabel(L["t"])
    plt.ylabel(L["tension"])
    plt.title(L["title_open_t"])
    _legend_below_center(ncol=3)
    plt.tight_layout(rect=[0, 0.10, 1, 1])
    p_open = os.path.join(plot_dir, f"open_tensions_{lang}.png")
    plt.savefig(p_open, dpi=180)
    plt.close()

    # 3) μ-时间（含 μss、μth、稳定段并集、tlife线）
    _cb(progress_cb, 78.0, "绘图：μ-时间（含 μss/μth/稳定段/tlife）...")
    tx, muf = _downsample_for_plot(t, mu_f, cfg.plot_max_points)
    plt.figure()
    plt.plot(tx, muf, label=L["mu_f"], color="tab:blue")

    if mu_ss is not None and np.isfinite(mu_ss):
        lbl = f'{L["mu_ss"]}={float(mu_ss):.4f}'
        plt.axhline(float(mu_ss), linestyle="--", label=lbl, color="tab:orange")
    if mu_th is not None and np.isfinite(mu_th):
        lbl = f'{L["mu_th"]}={float(mu_th):.4f}'
        plt.axhline(float(mu_th), linestyle="--", label=lbl, color="tab:red")

    # 连续稳定段窗口（并集，且已经“超限后停止判定”）
    if stable_segs_clip:
        first = True
        for k0, k1 in stable_segs_clip:
            t0 = float(t[k0])
            t1 = float(t[k1]) if (k1 < len(t)) else float(t[-1])
            if first:
                plt.axvspan(t0, t1, alpha=0.30, label=L["ss"], color="tab:cyan", zorder=1, linewidth=0)
                first = False
            else:
                plt.axvspan(t0, t1, alpha=0.30, color="tab:cyan", zorder=1, linewidth=0)

    # tlife：只在图例显示数值（不在图中显示文字）
    if tlife is not None:
        label = f'{L["tlife"]}≈{float(tlife):.1f}s'
        plt.axvline(float(tlife), linestyle="--", label=label, color="tab:purple")

    plt.xlabel(L["t"])
    plt.ylabel(L["mu"])
    plt.title(L["title_mu"])
    _legend_below_center(ncol=2)
    plt.tight_layout(rect=[0, 0.14, 1, 1])
    p_mu = os.path.join(plot_dir, f"mu_with_baseline_threshold_{lang}.png")
    plt.savefig(p_mu, dpi=180)
    plt.close()

    # 4) 开环 vs 闭环平均张力
    _cb(progress_cb, 92.0, "绘图：开环 vs 闭环平均张力...")

    # 对比输出范围：支持设置 [t_start, t_end]，(0,-1) 表示全部
    t0_cmp = float(getattr(cfg, "compare_t_start_s", 0.0))
    t1_cmp = float(getattr(cfg, "compare_t_end_s", -1.0))
    if (t0_cmp == 0.0) and (t1_cmp < 0.0):
        t_cmp = t
        op_cmp = op["tavg"]
        cl_cmp = cl["tavg"]
    else:
        if t1_cmp < 0.0:
            t1_cmp = float(t[-1])
        t0_cmp = max(0.0, t0_cmp)
        t1_cmp = max(t0_cmp, t1_cmp)
        mask = (t >= t0_cmp) & (t <= t1_cmp)
        if int(mask.sum()) < 2:
            t_cmp = t
            op_cmp = op["tavg"]
            cl_cmp = cl["tavg"]
        else:
            t_cmp = t[mask]
            op_cmp = op["tavg"][mask]
            cl_cmp = cl["tavg"][mask]

    tx, o = _downsample_for_plot(t_cmp, op_cmp, cfg.plot_max_points)
    _, c = _downsample_for_plot(t_cmp, cl_cmp, cfg.plot_max_points)
    plt.figure()
    plt.plot(tx, o, label=L["open"], color="tab:orange")
    plt.plot(tx, c, label=L["closed"], color="tab:blue")
    plt.xlabel(L["t"])
    plt.ylabel(L["tavg"])
    plt.title(L["title_cmp"])
    _legend_below_center(ncol=2)
    plt.tight_layout(rect=[0, 0.10, 1, 1])
    p_cmp = os.path.join(plot_dir, f"open_vs_closed_tavg_{lang}.png")
    plt.savefig(p_cmp, dpi=180)
    plt.close()

    _cb(progress_cb, 100.0, "图片导出完成")
    return {
        "plot_closed_tensions": p_closed,
        "plot_open_tensions": p_open,
        "plot_mu": p_mu,
        "plot_open_vs_closed": p_cmp,
        "plot_dir": plot_dir,
    }


def export_summary(res: Dict[str, Any], out_dir: str, extra: Optional[Dict[str, Any]] = None) -> str:
    """导出 summary.json"""
    out_dir = os.path.abspath(out_dir)
    _ensure_dir(out_dir)
    cfg: SimConfig = res["cfg"]

    s = {
        "config": asdict(cfg),
        "derived": res["derived"],
        "baseline": {"mu_ss": res["baseline"].get("mu_ss"), "stable_segments_idx": res["baseline"].get("stable_segments_idx"), "tlife_idx": res["baseline"].get("tlife_idx")},
        "threshold": {"mu_th": res["threshold"].get("mu_th")},
        "life": res["life"],
        "seed": res["seed"],
        "notes": {"scipy_used": HAVE_SCIPY},
    }
    if extra:
        s.update(extra)
    path = os.path.join(out_dir, "summary.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(s, f, ensure_ascii=False, indent=2)
    return path


def run_simulation(cfg: SimConfig, seed: int, out_dir: str, plot_lang: str = "zh", mode: str = "both", progress_cb: ProgressCB = None) -> Dict[str, Any]:
    """一键运行：仿真 + 导出（CLI 用）"""
    _ensure_dir(out_dir)
    _cb(progress_cb, 1.0, "开始仿真...")
    res = simulate(cfg, seed=seed, progress_cb=progress_cb)

    outputs = {}
    if mode in ("both", "xlsx"):
        outputs["xlsx"] = export_xlsx(res, out_dir=out_dir, progress_cb=progress_cb)
    if mode in ("both", "plots"):
        outputs.update(export_plots(res, out_dir=out_dir, lang=plot_lang, progress_cb=progress_cb))

    export_summary(res, out_dir=out_dir, extra={"outputs": outputs})
    return {"res": res, "outputs": outputs}


def _cli():
    import argparse
    p = argparse.ArgumentParser(description="针钩磨损平台全过程仿真（核心引擎）")
    p.add_argument("--theta_deg", type=float, default=20.0, help="包角(度)（默认20）")
    p.add_argument("--t_set", type=float, default=5.0, help="平均张力设定(N)")
    p.add_argument("--fs", type=float, default=50.0, help="采样率(Hz)（仅生成时间轴）")
    p.add_argument("--duration_s", type=float, default=600.0, help="采样时间(s)（仅生成时间轴）")
    p.add_argument("--rpm", type=float, default=300.0, help="转速(rpm)（固定值；CLI 旧模式）")
    p.add_argument("--mech_harmonic", type=int, default=1, help="机械扰动倍频m（1=一次转频）")
    p.add_argument("--out_dir", type=str, default="sim_out", help="输出目录")
    p.add_argument("--seed", type=int, default=7, help="随机种子")
    p.add_argument("--mode", type=str, default="both", choices=["both", "xlsx", "plots"], help="导出模式")
    p.add_argument("--plot_lang", type=str, default="zh", choices=["zh", "en"], help="图片语言")
    args = p.parse_args()

    cfg = SimConfig(
        theta_deg=args.theta_deg,
        t_set_N=args.t_set,
        fs_Hz=args.fs,
        duration_s=args.duration_s,
        rpm=args.rpm,
        mech_harmonic=args.mech_harmonic,
    )
    # CLI 保持旧行为：rpm 固定值
    cfg.rpm_min = args.rpm
    cfg.rpm_max = args.rpm

    def cb(pct, msg):
        print(f"[{pct:6.2f}%] {msg}")

    run_simulation(cfg, seed=args.seed, out_dir=args.out_dir, plot_lang=args.plot_lang, mode=args.mode, progress_cb=cb)
    print("完成。输出目录：", os.path.abspath(args.out_dir))


if __name__ == "__main__":
    _cli()

    # 阶段过渡速度系数（用于调整“段间过渡速度”，保证连续）
    # - 磨合→稳定：系数越大，衰减越快（过渡越快）
    # - 稳定→加速：系数越大，S 形上升越陡（过渡越快）
    trans_runin2stable_k: float = 1.0
    trans_stable2severe_k: float = 1.0

    # 开环-闭环对比输出范围（单位：s）
    # - (0, -1) 表示全部输出（默认）
    compare_t_start_s: float = 0.0
    compare_t_end_s: float = -1.0
