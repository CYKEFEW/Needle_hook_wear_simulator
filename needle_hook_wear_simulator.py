# -*- coding: utf-8 -*-
"""
针钩磨损平台检测全过程仿真程序（Capstan 反演 + 扰动 + 滤波 + 开环/闭环对比）
- UTF-8 编码
- 注释为中文
- 输出：xlsx（自动分 sheet）、png 曲线图、summary.json

用法示例：
python needle_hook_wear_simulator.py --theta_deg 180 --t_set 5 --fs 50 --duration_s 36000 --out_dir sim_out

依赖：
pip install numpy pandas openpyxl matplotlib
"""

from __future__ import annotations

import argparse
import json
import math
import os
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Tuple, Optional, List

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

# 尝试设置中文字体（不同系统可用字体不同，按顺序尝试；若都不可用也不会影响计算）
plt.rcParams['font.sans-serif'] = [
    'SimHei', 'Microsoft YaHei', 'Noto Sans CJK SC', 'Arial Unicode MS', 'DejaVu Sans'
]
plt.rcParams['axes.unicode_minus'] = False


# =========================
# 1) 基础工具：IIR 滤波、陷波、Hampel
# =========================

class IIR1LowPass:
    """一阶低通：y[n] = y[n-1] + alpha * (x[n] - y[n-1])
    其中 alpha = dt / (RC + dt), RC = 1/(2*pi*fc)
    """
    def __init__(self, fc_hz: float, fs_hz: float):
        self.fc = max(1e-6, float(fc_hz))
        self.fs = float(fs_hz)
        dt = 1.0 / self.fs
        rc = 1.0 / (2.0 * math.pi * self.fc)
        self.alpha = dt / (rc + dt)
        self.y = None

    def reset(self, y0: float = 0.0):
        self.y = float(y0)

    def step(self, x: float) -> float:
        if self.y is None or math.isnan(self.y):
            self.y = float(x)
            return self.y
        self.y = self.y + self.alpha * (x - self.y)
        return self.y


class NotchBiquad:
    """二阶陷波（notch）biquad，适合抑制机械周期扰动主频
    参考常用数字 biquad 形式（RBJ cookbook）
    """
    def __init__(self, f0_hz: float, fs_hz: float, q: float = 15.0):
        self.f0 = float(f0_hz)
        self.fs = float(fs_hz)
        self.q = max(0.5, float(q))
        self._design()
        self.x1 = self.x2 = 0.0
        self.y1 = self.y2 = 0.0

    def _design(self):
        w0 = 2.0 * math.pi * self.f0 / self.fs
        alpha = math.sin(w0) / (2.0 * self.q)

        b0 = 1.0
        b1 = -2.0 * math.cos(w0)
        b2 = 1.0
        a0 = 1.0 + alpha
        a1 = -2.0 * math.cos(w0)
        a2 = 1.0 - alpha

        # 归一化
        self.b0 = b0 / a0
        self.b1 = b1 / a0
        self.b2 = b2 / a0
        self.a1 = a1 / a0
        self.a2 = a2 / a0

    def reset(self):
        self.x1 = self.x2 = 0.0
        self.y1 = self.y2 = 0.0

    def step(self, x: float) -> float:
        y = self.b0 * x + self.b1 * self.x1 + self.b2 * self.x2 - self.a1 * self.y1 - self.a2 * self.y2
        self.x2, self.x1 = self.x1, x
        self.y2, self.y1 = self.y1, y
        return y


def hampel_filter_nan(x: np.ndarray, window: int = 25, n_sigma: float = 3.0) -> np.ndarray:
    """Hampel 异常点滤波：将异常点置为 NaN（不做“修复”，避免引入伪平滑）
    window: 半窗长（总窗长=2*window+1）
    """
    x = x.astype(float).copy()
    n = len(x)
    if n == 0:
        return x
    k = int(window)
    if k < 1:
        return x

    for i in range(n):
        lo = max(0, i - k)
        hi = min(n, i + k + 1)
        w = x[lo:hi]
        w = w[~np.isnan(w)]
        if len(w) < 5:
            continue
        med = np.median(w)
        mad = np.median(np.abs(w - med)) + 1e-12
        sigma = 1.4826 * mad
        if np.abs(x[i] - med) > n_sigma * sigma:
            x[i] = np.nan
    return x


# =========================
# 2) 仿真配置
# =========================

@dataclass
class SimConfig:
    # 用户输入核心参数
    theta_deg: float = 180.0          # 包角（度）
    t_set: float = 5.0               # 平均张力设定（N）
    fs: float = 50.0                 # 采样率（Hz）
    duration_s: float = 3600.0       # 采样时间（秒）

    # 扰动与噪声
    hf_noise_std: float = 0.12       # 高频噪声标准差（N）
    mech_freq_hz: float = 1.2        # 机械周期扰动主频（Hz）
    mech_amp_open: float = 0.60      # 开环周期扰动幅值（N）
    mech_amp_closed: float = 0.20    # 闭环残余周期扰动幅值（N）
    drift_rw_std: float = 0.002      # 随机游走漂移（N/样本）

    # 传感器噪声（紧边/松边额外独立噪声）
    sensor_std: float = 0.08         # N

    # 闭环控制（平均张力）
    plant_tau_s: float = 0.25        # 张力执行对象时间常数（越小越“快”）
    pid_kp: float = 3.0
    pid_ki: float = 1.2
    pid_kd: float = 0.0
    u_limit: float = 20.0            # 控制输入限幅（N，等效）

    # 预处理滤波（用于张力反馈/展示）
    t_lowpass_fc_hz: float = 6.0     # 张力反馈低通截止
    notch_enable: bool = True        # 是否启用陷波
    notch_q: float = 18.0            # 陷波 Q（越大越窄）

    # Capstan 反演门控与裁剪
    t_min_gate: float = 0.8          # 张力有效门限（N）
    ratio_clip: float = 50.0         # 张力比值裁剪上限（防止 log 发散）

    # μ 估计滤波
    mu_lowpass_fc_hz: float = 1.0    # μ 曲线低通（用于展示稳定趋势）
    hampel_window: int = 25
    hampel_n_sigma: float = 3.0

    # 稳定段基线提取（窗口 + 离散度 + 斜率）
    stable_window_s: float = 120.0   # 稳定段窗口时长（秒）
    stable_sigma_max: float = 0.010  # μ 的标准差阈值
    stable_slope_max: float = 1e-4   # μ 的斜率阈值（绝对值，μ/s）
    stable_hold_s: float = 60.0      # 连续满足条件的最短持续时间（秒）

    # 失效阈值（相对基线）与持续超限时间
    fail_delta_rel: float = 0.30     # μ_thr = μ0*(1+delta)
    fail_hold_s: float = 30.0        # 连续超限持续时间（秒）判为失效

    # 输出
    out_dir: str = "sim_out"
    seed: int = 7


# =========================
# 3) 摩擦系数“真值”曲线（磨合→稳定→加速）
# =========================

def mu_profile(t: np.ndarray, total_s: float) -> np.ndarray:
    """给一个工程上合理的 μ(t) 真值曲线（可按你的实验统计替换）
    - 早期磨合：μ 从高值指数回落到稳定区
    - 稳定磨损：缓慢漂移
    - 加速磨损：后期上升（模拟表面恶化/毛羽增多/局部刮擦等）
    """
    t = np.asarray(t, dtype=float)
    T = float(total_s)
    # 基线水平（可理解为稳定段平均摩擦系数）
    mu0 = 0.22

    # 三段时间比例
    t1 = 0.12 * T   # 磨合结束
    t2 = 0.78 * T   # 稳定段结束

    mu = np.empty_like(t)

    # 1) 磨合回落
    idx1 = t <= t1
    # 从 mu0+0.12 衰减到 mu0
    mu[idx1] = mu0 + 0.12 * np.exp(-t[idx1] / max(1.0, 0.25 * t1))

    # 2) 稳定段（轻微漂移）
    idx2 = (t > t1) & (t <= t2)
    drift = 0.01 * (t[idx2] - t1) / max(1.0, (t2 - t1))
    mu[idx2] = mu0 + drift

    # 3) 加速段（上升更快：用 sigmoid）
    idx3 = t > t2
    x = (t[idx3] - t2) / max(1.0, (T - t2))
    rise = 0.14 / (1.0 + np.exp(-10.0 * (x - 0.35)))  # 0→0.14
    mu[idx3] = (mu0 + 0.01) + rise

    return mu


# =========================
# 4) 从平均张力 + μ 真值生成紧边/松边（Capstan）
# =========================

def capstan_split_from_tavg_mu(t_avg: np.ndarray, mu: np.ndarray, theta_rad: float) -> Tuple[np.ndarray, np.ndarray]:
    """给定平均张力 Tavg=(Th+Tl)/2 与 μ，按 Capstan 比值 Th/Tl=exp(μθ) 反解 Th、Tl"""
    t_avg = np.asarray(t_avg, dtype=float)
    mu = np.asarray(mu, dtype=float)
    ratio = np.exp(mu * theta_rad)
    # 避免数值爆炸
    ratio = np.clip(ratio, 1.0 + 1e-9, 1e9)
    th = 2.0 * t_avg * ratio / (ratio + 1.0)
    tl = 2.0 * t_avg / (ratio + 1.0)
    return th, tl


# =========================
# 5) 张力仿真：开环与闭环
# =========================

def simulate_open_loop_tavg(cfg: SimConfig, t: np.ndarray, rng: np.random.Generator) -> np.ndarray:
    """开环：明显周期扰动 + 高频噪声 + 漂移"""
    dt = 1.0 / cfg.fs
    mech = cfg.mech_amp_open * np.sin(2.0 * np.pi * cfg.mech_freq_hz * t)
    hf = rng.normal(0.0, cfg.hf_noise_std, size=len(t))
    # 随机游走漂移
    rw = np.cumsum(rng.normal(0.0, cfg.drift_rw_std, size=len(t)))
    return cfg.t_set + mech + hf + rw


def simulate_closed_loop_tavg(cfg: SimConfig, t: np.ndarray, rng: np.random.Generator) -> Tuple[np.ndarray, np.ndarray, np.ndarray]:
    """闭环：PID 控平均张力，执行对象为一阶惯性 + 扰动注入
    返回：
    - t_avg_true：真实平均张力（闭环后）
    - t_avg_meas：测得平均张力（含测量噪声）
    - t_avg_filt：用于控制与展示的滤波后张力
    """
    fs = cfg.fs
    dt = 1.0 / fs

    # 扰动：残余周期扰动 + 高频噪声 + 漂移
    mech = cfg.mech_amp_closed * np.sin(2.0 * np.pi * cfg.mech_freq_hz * t)
    hf = rng.normal(0.0, cfg.hf_noise_std, size=len(t))
    rw = np.cumsum(rng.normal(0.0, cfg.drift_rw_std, size=len(t)))

    # 反馈测量滤波器
    lp = IIR1LowPass(cfg.t_lowpass_fc_hz, fs)
    lp.reset(cfg.t_set)
    notch = NotchBiquad(cfg.mech_freq_hz, fs, q=cfg.notch_q) if cfg.notch_enable else None
    if notch:
        notch.reset()

    # PID 状态
    integ = 0.0
    prev_e = 0.0

    # 执行对象状态（平均张力）
    T = cfg.t_set

    t_true = np.zeros_like(t, dtype=float)
    t_meas = np.zeros_like(t, dtype=float)
    t_filt = np.zeros_like(t, dtype=float)

    for i in range(len(t)):
        # 测量：真实 + 扰动（物理扰动）+ 传感噪声
        T_dist = T + mech[i] + hf[i] + rw[i]
        y_meas = T_dist + rng.normal(0.0, cfg.sensor_std)

        # notch + 低通
        y_nf = notch.step(y_meas) if notch else y_meas
        y_f = lp.step(y_nf)

        # PID
        e = cfg.t_set - y_f
        integ = integ + e * dt
        deriv = (e - prev_e) / dt
        prev_e = e

        u = cfg.pid_kp * e + cfg.pid_ki * integ + cfg.pid_kd * deriv
        u = float(np.clip(u, -cfg.u_limit, cfg.u_limit))

        # 一阶对象：dT/dt = (u - T)/tau
        tau = max(1e-3, cfg.plant_tau_s)
        T = T + dt * ((u - T) / tau)

        t_true[i] = T_dist
        t_meas[i] = y_meas
        t_filt[i] = y_f

    return t_true, t_meas, t_filt


# =========================
# 6) μ 反演、基线与寿命判定
# =========================

def estimate_mu_from_tensions(th: np.ndarray, tl: np.ndarray, theta_rad: float, cfg: SimConfig) -> Tuple[np.ndarray, np.ndarray]:
    """由测得 Th/Tl 反演 μ，并输出有效标志 q_valid"""
    th = np.asarray(th, dtype=float)
    tl = np.asarray(tl, dtype=float)

    q = np.ones_like(th, dtype=int)

    # 门控：张力过小会导致相对误差放大（ratio 与 log 发散风险）
    invalid = (th < cfg.t_min_gate) | (tl < cfg.t_min_gate) | np.isnan(th) | np.isnan(tl)
    q[invalid] = 0

    ratio = np.full_like(th, np.nan, dtype=float)
    ratio[~invalid] = th[~invalid] / tl[~invalid]

    # 理论上紧边>松边，ratio>=1；工程上做裁剪防止离群导致 log 爆炸
    ratio = np.clip(ratio, 1.0 + 1e-9, cfg.ratio_clip)

    mu = np.full_like(th, np.nan, dtype=float)
    mu[~invalid] = np.log(ratio[~invalid]) / max(1e-12, theta_rad)

    return mu, q


def find_stable_baseline(mu_f: np.ndarray, q: np.ndarray, t: np.ndarray, cfg: SimConfig) -> Tuple[Optional[float], Optional[float]]:
    """稳定段基线提取：
    - 滑动窗口
    - 有效样本比例 >= 0.9
    - 标准差 <= sigma_max
    - 线性拟合斜率 <= slope_max
    - 连续满足 stable_hold_s 后锁定 baseline
    返回：mu0, t0
    """
    fs = cfg.fs
    N = len(mu_f)
    win = int(round(cfg.stable_window_s * fs))
    hold = int(round(cfg.stable_hold_s * fs))
    if win < 20 or N < win + hold:
        return None, None

    # 预先计算方便
    ok_run = 0
    for start in range(0, N - win):
        seg = mu_f[start:start+win]
        seg_q = q[start:start+win]
        # 有效样本
        valid = (seg_q == 1) & (~np.isnan(seg))
        if valid.sum() < 0.9 * win:
            ok_run = 0
            continue
        y = seg[valid]
        x = t[start:start+win][valid]

        sigma = float(np.std(y))
        if sigma > cfg.stable_sigma_max:
            ok_run = 0
            continue

        # 斜率：用最小二乘拟合 y = a*x + b
        x0 = x - x.mean()
        denom = float(np.sum(x0 * x0)) + 1e-12
        a = float(np.sum(x0 * (y - y.mean())) / denom)

        if abs(a) > cfg.stable_slope_max:
            ok_run = 0
            continue

        ok_run += 1
        if ok_run >= max(1, hold // 10):  # 为降低计算，按“连续窗口”近似累计
            # baseline 用窗口中位数更稳健
            mu0 = float(np.nanmedian(seg[valid]))
            t0 = float(t[start])
            return mu0, t0

    return None, None


def detect_failure_time(mu_f: np.ndarray, q: np.ndarray, t: np.ndarray, mu0: float, cfg: SimConfig) -> Tuple[float, Optional[float]]:
    """失效判据：μ 超过阈值 μ_thr = μ0*(1+delta) 且持续 fail_hold_s"""
    fs = cfg.fs
    hold_n = int(round(cfg.fail_hold_s * fs))
    mu_thr = float(mu0 * (1.0 + cfg.fail_delta_rel))

    above = (q == 1) & (~np.isnan(mu_f)) & (mu_f > mu_thr)
    # 连续超限 run-length
    run = 0
    for i in range(len(above)):
        if above[i]:
            run += 1
            if run >= max(1, hold_n):
                return mu_thr, float(t[i - hold_n + 1])
        else:
            run = 0
    return mu_thr, None


# =========================
# 7) Excel 导出（自动分 sheet）
# =========================

EXCEL_SHEET_MAX_ROWS = 1_048_576  # Excel 单 sheet 极限

def write_df_chunked(writer: pd.ExcelWriter, df: pd.DataFrame, base_sheet: str):
    """若 df 超过单 sheet 行数，自动拆分多个 sheet 写入同一个 xlsx"""
    max_rows = EXCEL_SHEET_MAX_ROWS - 1  # 预留表头
    n = len(df)
    if n <= max_rows:
        df.to_excel(writer, sheet_name=base_sheet, index=False)
        return
    parts = int(math.ceil(n / max_rows))
    for i in range(parts):
        sl = df.iloc[i*max_rows:(i+1)*max_rows].copy()
        sl.to_excel(writer, sheet_name=f"{base_sheet}_{i+1}", index=False)


# =========================
# 8) 主流程：仿真 + 输出
# =========================

def run_simulation(cfg: SimConfig, seed: Optional[int] = None, out_dir: Optional[str] = None):
    """执行仿真并输出文件，返回 closed_df, open_df, baseline_info, life_info"""
    if seed is None:
        seed = cfg.seed
    rng = np.random.default_rng(int(seed))

    if out_dir is None:
        out_dir = cfg.out_dir
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    (out_dir / "plots").mkdir(parents=True, exist_ok=True)

    fs = float(cfg.fs)
    dt = 1.0 / fs
    n = int(round(cfg.duration_s * fs))
    if n < 10:
        raise ValueError("采样时间太短，样本数不足。")
    t = np.arange(n, dtype=float) * dt

    theta_rad = float(cfg.theta_deg) * math.pi / 180.0

    # 真值 μ
    mu_true = mu_profile(t, cfg.duration_s)

    # 1) 开环平均张力
    tavg_open = simulate_open_loop_tavg(cfg, t, rng)
    th_open_true, tl_open_true = capstan_split_from_tavg_mu(tavg_open, mu_true, theta_rad)

    # 加上传感器噪声（紧边/松边独立噪声）
    th_open_meas = th_open_true + rng.normal(0.0, cfg.sensor_std, size=n)
    tl_open_meas = tl_open_true + rng.normal(0.0, cfg.sensor_std, size=n)

    mu_open_raw, q_open = estimate_mu_from_tensions(th_open_meas, tl_open_meas, theta_rad, cfg)
    mu_open_hamp = hampel_filter_nan(mu_open_raw, window=cfg.hampel_window, n_sigma=cfg.hampel_n_sigma)

    mu_lp_open = IIR1LowPass(cfg.mu_lowpass_fc_hz, fs)
    mu_open_f = np.full_like(mu_open_hamp, np.nan, dtype=float)
    mu_lp_open.reset(float(np.nanmean(mu_open_hamp[np.isfinite(mu_open_hamp)]) if np.any(np.isfinite(mu_open_hamp)) else 0.2))
    for i in range(n):
        if np.isnan(mu_open_hamp[i]):
            mu_open_f[i] = np.nan
        else:
            mu_open_f[i] = mu_lp_open.step(float(mu_open_hamp[i]))

    # 2) 闭环平均张力
    tavg_closed_true, tavg_closed_meas, tavg_closed_filt = simulate_closed_loop_tavg(cfg, t, rng)
    th_closed_true, tl_closed_true = capstan_split_from_tavg_mu(tavg_closed_true, mu_true, theta_rad)

    th_closed_meas = th_closed_true + rng.normal(0.0, cfg.sensor_std, size=n)
    tl_closed_meas = tl_closed_true + rng.normal(0.0, cfg.sensor_std, size=n)

    mu_closed_raw, q_closed = estimate_mu_from_tensions(th_closed_meas, tl_closed_meas, theta_rad, cfg)
    mu_closed_hamp = hampel_filter_nan(mu_closed_raw, window=cfg.hampel_window, n_sigma=cfg.hampel_n_sigma)

    mu_lp_closed = IIR1LowPass(cfg.mu_lowpass_fc_hz, fs)
    mu_closed_f = np.full_like(mu_closed_hamp, np.nan, dtype=float)
    mu_lp_closed.reset(float(np.nanmean(mu_closed_hamp[np.isfinite(mu_closed_hamp)]) if np.any(np.isfinite(mu_closed_hamp)) else 0.2))
    for i in range(n):
        if np.isnan(mu_closed_hamp[i]):
            mu_closed_f[i] = np.nan
        else:
            mu_closed_f[i] = mu_lp_closed.step(float(mu_closed_hamp[i]))

    # 3) 基线与失效
    mu0, t0 = find_stable_baseline(mu_closed_f, q_closed, t, cfg)
    if mu0 is None:
        # 兜底：用前 10% 有效数据的中位数当 baseline（但会在 summary 里标注）
        valid = (q_closed == 1) & np.isfinite(mu_closed_f)
        cut = int(0.1 * n)
        mu0 = float(np.nanmedian(mu_closed_f[:cut][valid[:cut]]) if np.any(valid[:cut]) else 0.22)
        t0 = float(t[0])

    mu_thr, life_t = detect_failure_time(mu_closed_f, q_closed, t, mu0, cfg)

    # 4) 组装 DataFrame
    def build_df(tag: str,
                 th: np.ndarray, tl: np.ndarray, tavg: np.ndarray,
                 mu_raw: np.ndarray, mu_hamp: np.ndarray, mu_f: np.ndarray,
                 q: np.ndarray,
                 tavg_meas: Optional[np.ndarray] = None,
                 tavg_filt: Optional[np.ndarray] = None) -> pd.DataFrame:
        ff = th - tl
        df = pd.DataFrame({
            "t_s": t,
            "t_high_N": th,
            "t_low_N": tl,
            "t_avg_N": tavg,
            "f_fric_N": ff,
            "mu_est_raw": mu_raw,
            "mu_est_hampel": mu_hamp,
            "mu_est_filt": mu_f,
            "mu_true": mu_true,
            "q_valid": q.astype(int),
        })
        if tavg_meas is not None:
            df["t_avg_meas_N"] = tavg_meas
        if tavg_filt is not None:
            df["t_avg_filt_N"] = tavg_filt
        return df

    open_df = build_df("open",
                       th_open_meas, tl_open_meas, tavg_open,
                       mu_open_raw, mu_open_hamp, mu_open_f,
                       q_open)

    closed_df = build_df("closed",
                         th_closed_meas, tl_closed_meas, tavg_closed_true,
                         mu_closed_raw, mu_closed_hamp, mu_closed_f,
                         q_closed,
                         tavg_meas=tavg_closed_meas,
                         tavg_filt=tavg_closed_filt)

    # 5) 导出 Excel
    xlsx_path = out_dir / "needle_hook_wear_sim.xlsx"
    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        write_df_chunked(writer, closed_df, "closed_loop")
        write_df_chunked(writer, open_df, "open_loop")

    # 6) 绘图
    # 6.1 闭环张力曲线
    plt.figure()
    plt.plot(t, closed_df["t_high_N"].to_numpy(), label="紧边张力 Th")
    plt.plot(t, closed_df["t_low_N"].to_numpy(), label="松边张力 Tl")
    plt.plot(t, closed_df["t_avg_N"].to_numpy(), label="平均张力 Tavg")
    plt.xlabel("时间 / s")
    plt.ylabel("张力 / N")
    plt.title("闭环：高/低侧张力与平均张力-时间曲线")
    plt.legend()
    plt.tight_layout()
    plt.savefig(out_dir / "plots" / "closed_tensions.png", dpi=160)
    plt.close()

    # 6.2 μ 曲线（含 baseline & threshold）
    plt.figure()
    plt.plot(t, closed_df["mu_est_filt"].to_numpy(), label="μ（滤波后）")
    plt.plot(t, closed_df["mu_true"].to_numpy(), label="μ 真值", alpha=0.6)
    plt.axhline(mu0, linestyle="--", label=f"稳定段基线 μ0={mu0:.3f}")
    plt.axhline(mu_thr, linestyle="--", label=f"失效阈值 μthr={mu_thr:.3f}")
    if life_t is not None:
        plt.axvline(life_t, linestyle="--", label=f"寿命 τ={life_t:.1f}s")
    plt.xlabel("时间 / s")
    plt.ylabel("摩擦系数 μ")
    plt.title("摩擦系数-时间曲线（含稳定段基线与失效阈值）")
    plt.legend()
    plt.tight_layout()
    plt.savefig(out_dir / "plots" / "mu_with_baseline_threshold.png", dpi=160)
    plt.close()

    # 6.3 开环 vs 闭环（平均张力）
    plt.figure()
    plt.plot(t, open_df["t_avg_N"].to_numpy(), label="开环 Tavg（扰动明显）")
    plt.plot(t, closed_df["t_avg_N"].to_numpy(), label="闭环 Tavg（更稳定）")
    plt.xlabel("时间 / s")
    plt.ylabel("平均张力 / N")
    plt.title("开环与闭环控制效果对比（平均张力-时间）")
    plt.legend()
    plt.tight_layout()
    plt.savefig(out_dir / "plots" / "open_vs_closed_tavg.png", dpi=160)
    plt.close()

    # 7) summary
    summary = {
        "config": asdict(cfg),
        "baseline": {"mu0": mu0, "t0_s": t0},
        "threshold": {"mu_thr": mu_thr, "delta_rel": cfg.fail_delta_rel},
        "life": {"tau_s": life_t, "hold_s": cfg.fail_hold_s},
        "outputs": {
            "xlsx": str(xlsx_path.name),
            "plots": ["plots/closed_tensions.png", "plots/mu_with_baseline_threshold.png", "plots/open_vs_closed_tavg.png"],
        },
    }
    with open(out_dir / "summary.json", "w", encoding="utf-8") as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)

    return closed_df, open_df, summary


def parse_args() -> SimConfig:
    p = argparse.ArgumentParser(description="针钩磨损平台检测全过程仿真（Capstan + 扰动 + 滤波 + 开环/闭环）")
    p.add_argument("--theta_deg", type=float, required=True, help="包角（度）")
    p.add_argument("--t_set", type=float, required=True, help="平均张力设定（N）")
    p.add_argument("--fs", type=float, required=True, help="采样率（Hz）")
    p.add_argument("--duration_s", type=float, required=True, help="采样时间（秒）")
    p.add_argument("--out_dir", type=str, default="sim_out", help="输出目录")
    p.add_argument("--seed", type=int, default=7, help="随机种子（复现实验用）")

    # 可选高级参数（不强制）
    p.add_argument("--mech_freq_hz", type=float, default=1.2, help="机械周期扰动主频（Hz）")
    p.add_argument("--notch_disable", action="store_true", help="禁用陷波滤波器")
    p.add_argument("--pid_kp", type=float, default=3.0, help="闭环 PID: Kp")
    p.add_argument("--pid_ki", type=float, default=1.2, help="闭环 PID: Ki")
    p.add_argument("--pid_kd", type=float, default=0.0, help="闭环 PID: Kd")

    args = p.parse_args()

    cfg = SimConfig(
        theta_deg=args.theta_deg,
        t_set=args.t_set,
        fs=args.fs,
        duration_s=args.duration_s,
        out_dir=args.out_dir,
        seed=args.seed,
        mech_freq_hz=args.mech_freq_hz,
        notch_enable=(not args.notch_disable),
        pid_kp=args.pid_kp,
        pid_ki=args.pid_ki,
        pid_kd=args.pid_kd,
    )
    return cfg


def main():
    cfg = parse_args()
    run_simulation(cfg, seed=cfg.seed, out_dir=cfg.out_dir)
    print("仿真完成。输出目录：", cfg.out_dir)
    print("Excel：needle_hook_wear_sim.xlsx")
    print("图片：plots/*.png")
    print("摘要：summary.json")


if __name__ == "__main__":
    main()
