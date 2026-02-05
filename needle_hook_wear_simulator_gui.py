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


def _setup_chinese_font(cb: ProgressCB = None) -> Dict[str, Any]:
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
        _cb(cb, 2.0, f"已设置绘图中文字体：{chosen}")
        return {"font_ok": True, "font_name": chosen}

    _cb(cb, 2.0, "未检测到常见中文字体：中文图可能仍会乱码（可安装微软雅黑/黑体/Noto Sans CJK等）")
    return {"font_ok": False, "font_name": None}


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


def _iir_notch(x: np.ndarray, fs: float, f0: float, q: float) -> np.ndarray:
    """陷波滤波：优先 SciPy iirnotch，否则频域窄带抑制（全局）"""
    if f0 <= 0 or f0 >= fs / 2:
        return x.copy()

    if HAVE_SCIPY:
        b, a = signal.iirnotch(w0=f0, Q=q, fs=fs)
        try:
            return signal.filtfilt(b, a, x, method="gust")
        except Exception:
            return signal.filtfilt(b, a, x)

    # 无 SciPy：频域抑制（用于仿真展示足够）
    X = np.fft.rfft(x)
    freqs = np.fft.rfftfreq(len(x), d=1.0 / fs)
    bw = max(0.05, f0 / max(1.0, q))
    mask = (freqs > (f0 - bw)) & (freqs < (f0 + bw))
    X[mask] = 0
    return np.fft.irfft(X, n=len(x))


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
    """写入 xlsx（同一文件多工作表，必要时自动拆分）"""
    total_rows = sum(len(df) for df in frames.values())
    total_rows = max(1, total_rows)
    written = 0

    def _p(msg: str):
        nonlocal written
        pct = base_pct + span_pct * (written / total_rows)
        _cb(progress_cb, min(99.0, pct), msg)

    wb = Workbook(write_only=True)
    try:
        if wb.worksheets:
            wb.remove(wb.worksheets[0])
    except Exception:
        pass

    for name, df in frames.items():
        cols = list(df.columns)
        n = len(df)
        if n == 0:
            ws = wb.create_sheet(title=(name[:31] or "sheet"))
            ws.append(cols)
            continue

        part, start = 1, 0
        while start < n:
            end = min(n, start + row_limit)
            sheet_name = name if (part == 1 and n <= row_limit) else f"{name}_{part}"
            ws = wb.create_sheet(title=sheet_name[:31])
            ws.append(cols)

            chunk = df.iloc[start:end].to_numpy(dtype=object)
            for i in range(chunk.shape[0]):
                ws.append(chunk[i].tolist())

            written += (end - start)
            _p(f"写入 xlsx：{name}（{end}/{n}）")
            start = end
            part += 1

    wb.save(xlsx_path)
    _cb(progress_cb, base_pct + span_pct, f"xlsx 写入完成：{os.path.basename(xlsx_path)}")


@dataclass
class SimConfig:
    # 核心输入：fs 与 duration 仅用于生成时间轴
    theta_deg: float = 20.0
    t_set_N: float = 5.0
    fs_Hz: float = 50.0
    duration_s: float = 600.0

    # 机械扰动：仅由 rpm 换算
    rpm: float = 300.0
    mech_harmonic: int = 1

    # 扰动（开环明显，闭环衰减）
    noise_rms_open: float = 0.25
    noise_rms_closed: float = 0.08
    mech_amp_open: float = 0.6
    mech_amp_closed: float = 0.12
    drift_amp_open: float = 0.8
    drift_amp_closed: float = 0.15
    drift_freq_hz: float = 0.01

    # 摩擦系数三阶段（真值）
    mu_runin_start: float = 0.35
    mu_stable: float = 0.25
    mu_severe_end: float = 0.55
    runin_ratio: float = 0.12
    severe_start_ratio: float = 0.82

    # 张力测量噪声
    sensor_rms: float = 0.06

    # 门控/裁剪（抑制“比值+对数”放大）
    tmin_gate_N: float = 0.5
    ratio_clip_min: float = 0.2
    ratio_clip_max: float = 30.0

    # 滤波
    hampel_win_s: float = 0.8
    hampel_nsig: float = 3.0
    lowpass_fc_hz: float = 2.0

    # 稳定段/寿命判据
    stable_win_s: float = 120.0
    stable_sigma_max: float = 0.05  # 默认 0.05
    stable_slope_max: float = 1e-4
    stable_valid_min: float = 0.9
    fail_delta: float = 0.25
    fail_hold_s: float = 30.0

    # 导出/绘图
    export_stride: int = 1
    plot_max_points: int = 200_000

    def validate(self) -> None:
        assert self.fs_Hz > 0
        assert self.duration_s > 0
        assert self.export_stride >= 1
        assert 0 < self.theta_deg < 1080
        assert self.mech_harmonic >= 1
        assert self.rpm > 0, "rpm 必须>0（仅使用转速输入机械主频）"

    def mech_freq(self) -> float:
        """机械主频（Hz）"""
        return (float(self.rpm) / 60.0) * float(self.mech_harmonic)

    def notch_q_used(self) -> float:
        """陷波 Q（自动）"""
        return _notch_q_from_rpm(self.rpm)


def _mu_profile(t: np.ndarray, cfg: SimConfig) -> np.ndarray:
    """μ(t) 真值：磨合回落 → 稳定 → 剧烈磨损上升"""
    T = t[-1] if len(t) else 0.0
    if T <= 0:
        return np.array([], dtype=float)

    t_runin = cfg.runin_ratio * T
    t_severe = cfg.severe_start_ratio * T

    mu = np.empty_like(t, dtype=float)

    idx1 = t <= t_runin
    if np.any(idx1):
        tau = max(1e-6, 0.25 * t_runin)
        mu[idx1] = cfg.mu_stable + (cfg.mu_runin_start - cfg.mu_stable) * np.exp(-t[idx1] / tau)

    idx2 = (t > t_runin) & (t <= t_severe)
    if np.any(idx2):
        span = max(1e-6, (t_severe - t_runin))
        drift = 0.02 * (t[idx2] - t_runin) / span
        mu[idx2] = cfg.mu_stable + drift

    idx3 = t > t_severe
    if np.any(idx3):
        span = max(1e-6, (T - t_severe))
        x = (t[idx3] - t_severe) / span
        s = 1.0 / (1.0 + np.exp(-10 * (x - 0.35)))
        mu[idx3] = (cfg.mu_stable + 0.02) + (cfg.mu_severe_end - (cfg.mu_stable + 0.02)) * s

    # 微弱高频成分（非机械主频）
    mu += 0.002 * np.sin(2 * np.pi * 0.2 * t)
    return mu


def _make_tavg_open_closed(t: np.ndarray, cfg: SimConfig, rng: np.random.Generator):
    """
    开环/闭环平均张力 T_avg(t)
    说明：不做逐步 PID 微步仿真，而是构造“开环=扰动明显，闭环=扰动衰减”的等效效果。
    """
    f_mech = cfg.mech_freq()

    drift_open = cfg.drift_amp_open * np.sin(2 * np.pi * cfg.drift_freq_hz * t + 0.3)
    mech_open = cfg.mech_amp_open * np.sin(2 * np.pi * f_mech * t) \
                + 0.3 * cfg.mech_amp_open * np.sin(2 * np.pi * (2 * f_mech) * t + 0.9)
    noise_open = rng.normal(0.0, cfg.noise_rms_open, size=len(t))
    t_open = cfg.t_set_N + drift_open + mech_open + noise_open

    drift_closed = cfg.drift_amp_closed * np.sin(2 * np.pi * cfg.drift_freq_hz * t + 0.3)
    mech_closed = cfg.mech_amp_closed * np.sin(2 * np.pi * f_mech * t) \
                  + 0.3 * cfg.mech_amp_closed * np.sin(2 * np.pi * (2 * f_mech) * t + 0.9)
    noise_closed = rng.normal(0.0, cfg.noise_rms_closed, size=len(t))
    residual = 0.02 * cfg.t_set_N * np.sin(2 * np.pi * (0.5 * f_mech) * t + 1.1)
    t_closed = cfg.t_set_N + drift_closed + mech_closed + noise_closed + residual

    return t_open, t_closed


def _tensions_from_tavg_mu(tavg: np.ndarray, mu: np.ndarray, cfg: SimConfig, rng: np.random.Generator):
    """
    由 T_avg 与 μ 生成紧边/松边张力：
    r = exp(μθ)，且 (T_high+T_low)/2 = T_avg
    => T_high = 2*T_avg*r/(1+r),  T_low = 2*T_avg/(1+r)
    """
    theta = np.deg2rad(cfg.theta_deg)
    r = np.exp(np.clip(mu * theta, -10, 10))
    t_high = 2.0 * tavg * r / (1.0 + r)
    t_low = 2.0 * tavg / (1.0 + r)

    t_high = np.maximum(t_high + rng.normal(0.0, cfg.sensor_rms, size=len(t_high)), 0.0)
    t_low = np.maximum(t_low + rng.normal(0.0, cfg.sensor_rms, size=len(t_low)), 0.0)
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
    稳定段基线 μss：窗口内有效比例 + std + slope
    end_idx：仅在 [0,end_idx) 范围内判定（用于“超限后不再判定稳定段”）
    """
    n_all = len(mu_f)
    if n_all == 0:
        return None, None
    n = n_all if (end_idx is None) else int(max(0, min(n_all, end_idx)))
    if n <= 0:
        return None, None

    win = int(round(cfg.stable_win_s * cfg.fs_Hz))
    win = max(win, 50)
    step = max(1, win // 10)

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

        yy = mu_f[k0:k1]
        tt = t[k0:k1]
        m = np.isfinite(yy)
        if m.sum() < 10:
            continue
        A = np.vstack([tt[m], np.ones(m.sum())]).T
        slope = float(np.linalg.lstsq(A, yy[m], rcond=None)[0][0])
        if abs(slope) > cfg.stable_slope_max:
            continue

        mu_ss = float(np.nanmean(yy[m]))
        return (k0, k1), mu_ss

    return None, None


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

        yy = mu_f[k0:k1]
        tt = t[k0:k1]
        mm = np.isfinite(yy)
        if mm.sum() < 10:
            continue
        A = np.vstack([tt[mm], np.ones(mm.sum())]).T
        slope = float(np.linalg.lstsq(A, yy[mm], rcond=None)[0][0])
        if abs(slope) > cfg.stable_slope_max:
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
    return segs


def _find_failure_time(mu_f: np.ndarray, t: np.ndarray, mu_ss: Optional[float], cfg: SimConfig):
    """寿命判据：μ 持续超过 μth=μss*(1+δ) 持续 W_hold，输出 tlife"""
    if mu_ss is None or not np.isfinite(mu_ss):
        return None, None
    mu_th = float(mu_ss * (1.0 + cfg.fail_delta))
    hold = int(round(cfg.fail_hold_s * cfg.fs_Hz))
    hold = max(1, hold)

    above = np.isfinite(mu_f) & (mu_f > mu_th)
    count = 0
    for i, a in enumerate(above):
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
    mu_true = _mu_profile(t, cfg)

    _cb(progress_cb, 18.0, "生成开环/闭环平均张力...")
    tavg_open, tavg_closed = _make_tavg_open_closed(t, cfg, rng)

    _cb(progress_cb, 24.0, "生成高/低张力侧张力（含传感器噪声）...")
    th_open, tl_open = _tensions_from_tavg_mu(tavg_open, mu_true, cfg, rng)
    th_cl, tl_cl = _tensions_from_tavg_mu(tavg_closed, mu_true, cfg, rng)

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

    f_mech = cfg.mech_freq()
    q_notch = cfg.notch_q_used()
    _cb(progress_cb, 40.0, f"μ 陷波（f_mech={f_mech:.6g}Hz, Q={q_notch:.3g}）...")
    mu_cl_notch = _iir_notch(mu_f0, fs=cfg.fs_Hz, f0=f_mech, q=q_notch)

    _cb(progress_cb, 44.0, "μ 低通滤波...")
    mu_cl_lp = _lowpass(mu_cl_notch, fs=cfg.fs_Hz, fc=cfg.lowpass_fc_hz, order=3).astype(float)
    mu_cl_lp[q_cl == 0] = np.nan

    _cb(progress_cb, 48.0, "稳定段基线与寿命判据计算...")
    # 1) 先识别“全时域”的连续稳定段（用于可视化），后续绘图会在 tlife 右侧裁剪掉
    stable_segs_all = _find_stable_segments(mu_cl_lp, t, q_cl, cfg, end_idx=None)

    # 2) 计算基线 μss：优先在“失效前”段落内求；若尚未失效，则使用全时域
    #    先粗略用全时域求 μss，再求 tlife；再把基线限制在 tlife 之前重新求一次（更符合工程意义）
    stable_idx0, mu_ss0 = _find_stable_baseline(mu_cl_lp, t, q_cl, cfg, end_idx=None)
    tlife, mu_th = _find_failure_time(mu_cl_lp, t, mu_ss0, cfg)

    tlife_idx = None
    if tlife is not None:
        tlife_idx = int(max(0, min(len(t) - 1, round(float(tlife) * cfg.fs_Hz))))
        stable_idx, mu_ss = _find_stable_baseline(mu_cl_lp, t, q_cl, cfg, end_idx=tlife_idx)
        # 若失效前找不到稳定段，则回退到全时域
        if mu_ss is None or not np.isfinite(mu_ss):
            stable_idx, mu_ss = stable_idx0, mu_ss0
    else:
        stable_idx, mu_ss = stable_idx0, mu_ss0

    # 3) 用最终 μss 更新 μth（保持一致）
    if mu_ss is not None and np.isfinite(mu_ss):
        mu_th = float(mu_ss * (1.0 + cfg.fail_delta))

    stable_segs = stable_segs_all

    res = {
        "cfg": cfg,
        "seed": int(seed),
        "derived": {
            "mech_freq_hz_used": float(f_mech),
            "notch_q_used": float(q_notch),
        },
        "series": {
            "t": t,
            "mu_true": mu_true,
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

    cl = res["series"]["closed"]
    op = res["series"]["open"]

    closed_df = pd.DataFrame({
        "t_s": t_e,
        "t_high_N": ds(cl["th"]),
        "t_low_N": ds(cl["tl"]),
        "t_avg_N": ds(cl["tavg"]),
        "f_fric_N": ds(cl["ff"]),
        "mu_true": ds(res["series"]["mu_true"]),
        "mu_raw": ds(cl["mu_raw"]),
        "mu_hampel": ds(cl["mu_hampel"]),
        "mu_notch": ds(cl["mu_notch"]),
        "mu_filt": ds(cl["mu_filt"]),
        "q_valid": ds(cl["q"]).astype(int),
        "rpm": np.full_like(t_e, float(cfg.rpm), dtype=float),
        "mech_freq_hz": np.full_like(t_e, f_mech, dtype=float),
        "notch_q_used": np.full_like(t_e, q_notch, dtype=float),
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
        "mu_raw": ds(op["mu_raw"]),
        "q_valid": ds(op["q"]).astype(int),
        "rpm": np.full_like(t_e, float(cfg.rpm), dtype=float),
        "mech_freq_hz": np.full_like(t_e, f_mech, dtype=float),
        "notch_q_used": np.full_like(t_e, q_notch, dtype=float),
        "theta_deg": np.full_like(t_e, float(cfg.theta_deg), dtype=float),
        "t_set_N": np.full_like(t_e, float(cfg.t_set_N), dtype=float),
    })

    xlsx_path = os.path.join(out_dir, "needle_hook_wear_sim.xlsx")
    _cb(progress_cb, 55.0, "开始写入 xlsx（多sheet，必要时自动拆分）...")
    _write_xlsx_multisheet(
        xlsx_path,
        {"closed_loop": closed_df, "open_loop": open_df},
        progress_cb=progress_cb,
        base_pct=55.0,
        span_pct=44.0
    )
    _cb(progress_cb, 100.0, "xlsx 导出完成")
    return xlsx_path


def _legend_below_center(ncol: int = 3):
    """图例放到下方居中（不在图框内）"""
    plt.legend(loc="upper center", bbox_to_anchor=(0.5, -0.18), ncol=ncol, frameon=False)


def export_plots(res: Dict[str, Any], out_dir: str, lang: str = "zh", progress_cb: ProgressCB = None) -> Dict[str, str]:
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

    if lang.lower().startswith("zh"):
        _cb(progress_cb, 52.0, "设置中文绘图字体...")
        _setup_chinese_font(progress_cb)
    else:
        matplotlib.rcParams["axes.unicode_minus"] = False

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
    p1 = os.path.join(plot_dir, f"closed_tensions_{lang}.png")
    plt.savefig(p1, dpi=180)
    plt.close()

    # 2) μ-时间（含 μss、μth、稳定段并集、tlife线）
    _cb(progress_cb, 75.0, "绘图：μ-时间（含 μss/μth/稳定段/tlife）...")
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
    p2 = os.path.join(plot_dir, f"mu_with_baseline_threshold_{lang}.png")
    plt.savefig(p2, dpi=180)
    plt.close()

    # 3) 开环 vs 闭环平均张力
    _cb(progress_cb, 90.0, "绘图：开环 vs 闭环平均张力...")
    tx, o = _downsample_for_plot(t, op["tavg"], cfg.plot_max_points)
    _, c = _downsample_for_plot(t, cl["tavg"], cfg.plot_max_points)
    plt.figure()
    plt.plot(tx, o, label=L["open"], color="tab:orange")
    plt.plot(tx, c, label=L["closed"], color="tab:blue")
    plt.xlabel(L["t"])
    plt.ylabel(L["tavg"])
    plt.title(L["title_cmp"])
    _legend_below_center(ncol=2)
    plt.tight_layout(rect=[0, 0.10, 1, 1])
    p3 = os.path.join(plot_dir, f"open_vs_closed_tavg_{lang}.png")
    plt.savefig(p3, dpi=180)
    plt.close()

    _cb(progress_cb, 100.0, "图片导出完成")
    return {"plot_closed_tensions": p1, "plot_mu": p2, "plot_open_vs_closed": p3, "plot_dir": plot_dir}


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
    p.add_argument("--rpm", type=float, default=300.0, help="转速(rpm)（唯一输入主频方式）")
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

    def cb(pct, msg):
        print(f"[{pct:6.2f}%] {msg}")

    run_simulation(cfg, seed=args.seed, out_dir=args.out_dir, plot_lang=args.plot_lang, mode=args.mode, progress_cb=cb)
    print("完成。输出目录：", os.path.abspath(args.out_dir))


if __name__ == "__main__":
    _cli()
