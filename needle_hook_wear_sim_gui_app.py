# -*- coding: utf-8 -*-
"""
needle_hook_wear_sim_gui_app.py

针钩磨损平台全过程仿真（GUI）
- “导出 xlsx”和“导出 图片”两个按钮
- 导出图片可选中文/英文
- 进度条显示百分比
- 机械主频只允许转速输入：f_mech=rpm/60*m（默认300rpm）
- 陷波滤波器 Q 不手动输入：Q≈clamp(15,80,rpm/10)
- 稳定段基线：μss；超限阈值：μth；tlife 数值仅在图例显示
- 规则：第一次超限后不再判定稳定段窗口
"""

import os
import threading
import queue
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from needle_hook_wear_simulator_gui import SimConfig, simulate, export_xlsx, export_plots, export_summary, HAVE_SCIPY


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("针钩磨损检测全过程仿真（Capstan + 扰动 + 滤波 + 判据）")
        self.geometry("1080x760")
        self.minsize(980, 700)

        self.msg_q = queue.Queue()
        self.worker = None

        self.cfg = SimConfig()  # 默认包角20°，默认rpm=300
        self.out_dir = tk.StringVar(value=os.path.abspath("sim_out"))
        self.seed = tk.IntVar(value=7)

        self.plot_lang = tk.StringVar(value="zh")  # zh / en

        # 缓存：避免连续导出重复计算
        self._cache_key = None
        self._cache_res = None

        top = ttk.Frame(self, padding=10)
        top.pack(fill="x")
        info = ("说明：fs 与采样时间仅用于生成时间轴（输出点数）。\n"
                "机械主频：仅由 rpm 换算 f_mech=rpm/60*m（默认300rpm）。\n"
                "陷波 Q：根据 rpm 自动估算（Q≈clamp(15,80,rpm/10)）。\n"
                "μ-时间图：含稳定段基线 μss、超限阈值 μth，显示连续稳定段（并集），标注 tlife（仅图例显示数值）。\n"
                "规则：第一次超限后不再判定稳定段窗口。\n"
                f"SciPy：{'已检测到（陷波/低通更快更标准）' if HAVE_SCIPY else '未检测到（将使用近似实现）'}")
        ttk.Label(top, text=info).pack(anchor="w")

        main = ttk.Frame(self, padding=10)
        main.pack(fill="both", expand=True)

        left = ttk.Frame(main)
        left.pack(side="left", fill="both", expand=True, padx=(0, 10))
        right = ttk.Frame(main)
        right.pack(side="right", fill="y")

        self._build_form(left)
        self._build_run_panel(right)

        self.status = tk.StringVar(value="就绪")
        bar = ttk.Frame(self, padding=(10, 5))
        bar.pack(fill="x")
        ttk.Label(bar, textvariable=self.status).pack(side="left")

        self.after(120, self._poll_msgs)

    def _build_form(self, parent):
        nb = ttk.Notebook(parent)
        nb.pack(fill="both", expand=True)

        tab_core = ttk.Frame(nb, padding=10)
        tab_dist = ttk.Frame(nb, padding=10)
        tab_phase = ttk.Frame(nb, padding=10)
        tab_filter = ttk.Frame(nb, padding=10)
        tab_judge = ttk.Frame(nb, padding=10)
        tab_export = ttk.Frame(nb, padding=10)
        nb.add(tab_core, text="核心输入")
        nb.add(tab_dist, text="扰动参数")
        nb.add(tab_phase, text="阶段比例/μ范围")
        nb.add(tab_filter, text="滤波参数")
        nb.add(tab_judge, text="基线/阈值/寿命")
        nb.add(tab_export, text="导出/绘图")

        self._vars = {}
        self._types = {}

        # 核心输入
        self._add_entry(tab_core, "包角 θ (deg)（默认20）", "theta_deg")
        self._add_entry(tab_core, "平均张力设定 T_set (N)", "t_set_N")
        self._add_entry(tab_core, "采样率 fs (Hz) 仅生成时间轴", "fs_Hz")
        self._add_entry(tab_core, "采样时间 duration (h) 仅生成时间轴", "duration_h")


        # 阶段比例（滑块 + 可输入百分比 + 可锁定一个滑块，总和=100%）
        ttk.Label(tab_phase, text="阶段时间比例（滑块/输入；自动保持总和=100%）").grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 8))
        tab_phase.columnconfigure(1, weight=1)

        self._ratio_lock = False
        # 三段比例（单位：%）
        self.phase_runin_pct = tk.DoubleVar(value=float(self.cfg.phase_runin_ratio) * 100.0)
        self.phase_stable_pct = tk.DoubleVar(value=float(self.cfg.phase_stable_ratio) * 100.0)
        self.phase_severe_pct = tk.DoubleVar(value=float(self.cfg.phase_severe_ratio) * 100.0)

        # 允许锁定一个阶段（锁定后该滑块与输入框禁用）
        self.locked_phase = tk.StringVar(value="none")  # none/runin/stable/severe

        # 输入框变量（百分比数值，不带%）
        self._lbl_runin = tk.StringVar()
        self._lbl_stable = tk.StringVar()
        self._lbl_severe = tk.StringVar()
        self.phase_sum_label = tk.StringVar()

        # 锁定选择
        ttk.Label(tab_phase, text="锁定：").grid(row=1, column=0, sticky="w", pady=4)
        lock_box = ttk.Frame(tab_phase)
        lock_box.grid(row=1, column=1, columnspan=3, sticky="w", pady=4)
        ttk.Radiobutton(lock_box, text="不锁定", variable=self.locked_phase, value="none", command=self._on_lock_change).pack(side="left", padx=(0, 10))
        ttk.Radiobutton(lock_box, text="锁定磨合", variable=self.locked_phase, value="runin", command=self._on_lock_change).pack(side="left", padx=(0, 10))
        ttk.Radiobutton(lock_box, text="锁定稳定", variable=self.locked_phase, value="stable", command=self._on_lock_change).pack(side="left", padx=(0, 10))
        ttk.Radiobutton(lock_box, text="锁定加速", variable=self.locked_phase, value="severe", command=self._on_lock_change).pack(side="left", padx=(0, 10))

        # 三段：滑块 + 输入框
        ttk.Label(tab_phase, text="磨合阶段").grid(row=2, column=0, sticky="w", pady=4)
        self._scale_runin = ttk.Scale(tab_phase, from_=0.0, to=100.0, variable=self.phase_runin_pct,
                                      command=lambda _v: self._on_ratio_change("runin"))
        self._scale_runin.grid(row=2, column=1, sticky="we", pady=4, padx=(0, 8))
        self._entry_runin = ttk.Entry(tab_phase, textvariable=self._lbl_runin, width=8)
        self._entry_runin.grid(row=2, column=2, sticky="w")
        ttk.Label(tab_phase, text="%").grid(row=2, column=3, sticky="w", padx=(4, 0))
        self._entry_runin.bind("<Return>", lambda _e: self._on_ratio_entry("runin"))
        self._entry_runin.bind("<FocusOut>", lambda _e: self._on_ratio_entry("runin"))

        ttk.Label(tab_phase, text="稳定磨损阶段").grid(row=3, column=0, sticky="w", pady=4)
        self._scale_stable = ttk.Scale(tab_phase, from_=0.0, to=100.0, variable=self.phase_stable_pct,
                                       command=lambda _v: self._on_ratio_change("stable"))
        self._scale_stable.grid(row=3, column=1, sticky="we", pady=4, padx=(0, 8))
        self._entry_stable = ttk.Entry(tab_phase, textvariable=self._lbl_stable, width=8)
        self._entry_stable.grid(row=3, column=2, sticky="w")
        ttk.Label(tab_phase, text="%").grid(row=3, column=3, sticky="w", padx=(4, 0))
        self._entry_stable.bind("<Return>", lambda _e: self._on_ratio_entry("stable"))
        self._entry_stable.bind("<FocusOut>", lambda _e: self._on_ratio_entry("stable"))

        ttk.Label(tab_phase, text="加速磨损阶段").grid(row=4, column=0, sticky="w", pady=4)
        self._scale_severe = ttk.Scale(tab_phase, from_=0.0, to=100.0, variable=self.phase_severe_pct,
                                       command=lambda _v: self._on_ratio_change("severe"))
        self._scale_severe.grid(row=4, column=1, sticky="we", pady=4, padx=(0, 8))
        self._entry_severe = ttk.Entry(tab_phase, textvariable=self._lbl_severe, width=8)
        self._entry_severe.grid(row=4, column=2, sticky="w")
        ttk.Label(tab_phase, text="%").grid(row=4, column=3, sticky="w", padx=(4, 0))
        self._entry_severe.bind("<Return>", lambda _e: self._on_ratio_entry("severe"))
        self._entry_severe.bind("<FocusOut>", lambda _e: self._on_ratio_entry("severe"))

        ttk.Label(tab_phase, textvariable=self.phase_sum_label).grid(row=5, column=0, columnspan=4, sticky="w", pady=(4, 10))

        # μ 范围输入（min/max）
        ttk.Label(tab_phase, text="三阶段摩擦系数范围（min/max；若 min==max 即固定值）").grid(row=5, column=0, columnspan=3, sticky="w", pady=(0, 8))
        self._add_entry(tab_phase, "磨合段 μ范围：min", "mu_runin_min", row=6)
        self._add_entry(tab_phase, "磨合段 μ范围：max", "mu_runin_max", row=7)
        self._add_entry(tab_phase, "稳定段 μ范围：min", "mu_stable_min", row=8)
        self._add_entry(tab_phase, "稳定段 μ范围：max", "mu_stable_max", row=9)
        self._add_entry(tab_phase, "加速段 μ范围：min", "mu_severe_min", row=10)
        self._add_entry(tab_phase, "加速段 μ范围：max", "mu_severe_max", row=11)

        self._update_ratio_labels()
        self._on_lock_change()

        # 扰动：只允许 rpm 输入主频（删除直接输入主频）
        ttk.Label(tab_dist, text="机械周期扰动主频：f_mech = rpm/60*m（仅此方式输入）").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))
        self._add_entry(tab_dist, "转速 rpm（默认300）", "rpm", row=1)
        self._add_entry(tab_dist, "倍频 m（1=一次转频）", "mech_harmonic", row=2, is_int=True)

        self._add_entry(tab_dist, "开环周期幅值 A_mech_open (N)", "mech_amp_open", row=3)
        self._add_entry(tab_dist, "闭环周期幅值 A_mech_closed (N)", "mech_amp_closed", row=4)
        self._add_entry(tab_dist, "开环噪声 RMS_open (N)", "noise_rms_open", row=5)
        self._add_entry(tab_dist, "闭环噪声 RMS_closed (N)", "noise_rms_closed", row=6)
        self._add_entry(tab_dist, "漂移频率 f_drift (Hz)", "drift_freq_hz", row=7)
        self._add_entry(tab_dist, "开环漂移幅值 A_drift_open (N)", "drift_amp_open", row=8)
        self._add_entry(tab_dist, "闭环漂移幅值 A_drift_closed (N)", "drift_amp_closed", row=9)
        self._add_entry(tab_dist, "张力传感器噪声 RMS_sensor (N)", "sensor_rms", row=10)

        # 显示 f_mech 与 Notch Q（自动）
        self.f_mech_var = tk.StringVar(value="f_mech=— Hz")
        ttk.Label(tab_dist, textvariable=self.f_mech_var).grid(row=11, column=0, columnspan=2, sticky="w", pady=(10, 0))
        self._vars["rpm"].trace_add("write", lambda *_: self._update_f_mech_label())
        self._vars["mech_harmonic"].trace_add("write", lambda *_: self._update_f_mech_label())
        self._update_f_mech_label()

        # 滤波
        ttk.Label(tab_filter, text="Hampel：异常点剔除(置NaN)；Notch：抑制机械周期（Q由rpm自动估算）；Lowpass：抑制高频噪声").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))
        self._add_entry(tab_filter, "Hampel 窗口 (s)", "hampel_win_s", row=1)
        self._add_entry(tab_filter, "Hampel nσ", "hampel_nsig", row=2)
        self._add_entry(tab_filter, "低通截止 fc (Hz)", "lowpass_fc_hz", row=3)
        self._add_entry(tab_filter, "门控下限 T_min (N)", "tmin_gate_N", row=4)
        self._add_entry(tab_filter, "比值裁剪 r_min", "ratio_clip_min", row=5)
        self._add_entry(tab_filter, "比值裁剪 r_max", "ratio_clip_max", row=6)

        # 判据
        ttk.Label(tab_judge, text="稳定段基线：std + slope + 有效比例；寿命：阈值线持续超限（第一次超限后不再判定稳定段）").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))
        self._add_entry(tab_judge, "稳定窗口 W_ss (s)", "stable_win_s", row=1)
        self._add_entry(tab_judge, "稳定标准差 σ_max（默认0.05）", "stable_sigma_max", row=2)
        self._add_entry(tab_judge, "稳定斜率阈值 |dμ/dt|_max", "stable_slope_max", row=3)
        self._add_entry(tab_judge, "稳定有效比例 q_min", "stable_valid_min", row=4)
        self._add_entry(tab_judge, "失效阈值增量 δ（μth=μss*(1+δ)）", "fail_delta", row=5)
        self._add_entry(tab_judge, "持续超限 W_hold (s)", "fail_hold_s", row=6)

        # 导出/绘图
        ttk.Label(tab_export, text="duration 很长时 xlsx 会很大；可用 stride 抽样导出。绘图会自动降采样。").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))
        self._add_entry(tab_export, "导出步长 stride（1=全量）", "export_stride", row=1, is_int=True)
        self._add_entry(tab_export, "绘图最大点数", "plot_max_points", row=2, is_int=True)

    def _add_entry(self, parent, label, field, row=None, is_int=False):
        if row is None:
            row = parent.grid_size()[1]
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=(0, 10), pady=4)
        if field == "duration_h":
            var = tk.StringVar(value=str(getattr(self.cfg, "duration_s") / 3600.0))
        else:
            var = tk.StringVar(value=str(getattr(self.cfg, field)))
        ttk.Entry(parent, textvariable=var, width=24).grid(row=row, column=1, sticky="w", pady=4)
        self._vars[field] = var
        self._types[field] = int if is_int else float

    def _update_f_mech_label(self):
        try:
            rpm = float(self._vars["rpm"].get().strip() or 300.0)
            m = int(float(self._vars["mech_harmonic"].get().strip() or 1))
            f = (rpm / 60.0) * max(1, m)
            q = max(15.0, min(80.0, rpm / 10.0))
            self.f_mech_var.set(f"f_mech = rpm/60*m = {f:.6g} Hz，Notch Q≈{q:.3g}")
        except Exception:
            self.f_mech_var.set("f_mech=— Hz（输入有误）")


    def _update_ratio_labels(self):
        r1 = float(self.phase_runin_pct.get())
        r2 = float(self.phase_stable_pct.get())
        r3 = float(self.phase_severe_pct.get())
        s = max(1e-9, r1 + r2 + r3)
        r1, r2, r3 = (100.0 * r1 / s, 100.0 * r2 / s, 100.0 * r3 / s)
        # 输入框显示为纯数字（不带%），便于直接输入
        self._lbl_runin.set(f"{r1:.2f}")
        self._lbl_stable.set(f"{r2:.2f}")
        self._lbl_severe.set(f"{r3:.2f}")
        self.phase_sum_label.set(f"磨合/稳定/加速：{r1:.2f}% / {r2:.2f}% / {r3:.2f}%（Σ=100%）")

    def _on_ratio_change(self, changed: str):
        """保持三段比例和为 100%。支持锁定一个阶段；未锁定时尽量保留另外两段的相对比例。"""
        if getattr(self, "_ratio_lock", False):
            return
        self._ratio_lock = True
        try:
            lock = getattr(self, "locked_phase", None)
            lock = lock.get() if lock is not None else "none"
            r1 = float(self.phase_runin_pct.get())
            r2 = float(self.phase_stable_pct.get())
            r3 = float(self.phase_severe_pct.get())

            def clamp(x, lo=0.0, hi=100.0):
                try:
                    x = float(x)
                except Exception:
                    x = 0.0
                return max(lo, min(hi, x))

            r1 = clamp(r1)
            r2 = clamp(r2)
            r3 = clamp(r3)

            if lock not in ("runin", "stable", "severe"):
                lock = "none"

            if lock == "none":
                # 未锁定：保持 changed 不动，另外两段按原比例缩放到剩余值
                if changed == "runin":
                    rem = 100.0 - r1
                    tot = r2 + r3
                    if tot <= 1e-9:
                        r2 = rem * 0.5
                        r3 = rem - r2
                    else:
                        r2 = rem * (r2 / tot)
                        r3 = rem - r2
                elif changed == "stable":
                    rem = 100.0 - r2
                    tot = r1 + r3
                    if tot <= 1e-9:
                        r1 = rem * 0.5
                        r3 = rem - r1
                    else:
                        r1 = rem * (r1 / tot)
                        r3 = rem - r1
                else:  # severe
                    rem = 100.0 - r3
                    tot = r1 + r2
                    if tot <= 1e-9:
                        r1 = rem * 0.5
                        r2 = rem - r1
                    else:
                        r1 = rem * (r1 / tot)
                        r2 = rem - r1
                # 归一化到 100
                s = max(1e-9, r1 + r2 + r3)
                r1, r2, r3 = (100.0 * r1 / s, 100.0 * r2 / s, 100.0 * r3 / s)
            else:
                # 锁定：锁定的那段固定，其余两段保持和为剩余值；改变一个，另一个自动补齐
                if lock == "runin":
                    L = clamp(r1)
                    rem = max(0.0, 100.0 - L)
                    if changed == "stable":
                        r2 = clamp(r2, 0.0, rem)
                        r3 = rem - r2
                    elif changed == "severe":
                        r3 = clamp(r3, 0.0, rem)
                        r2 = rem - r3
                    else:
                        # 理论上不会发生（锁定的控件会被禁用），这里做保护
                        # 保留 r2,r3 比例并归一化到 rem
                        tot = r2 + r3
                        if tot <= 1e-9:
                            r2 = rem * 0.5
                            r3 = rem - r2
                        else:
                            r2 = rem * (r2 / tot)
                            r3 = rem - r2
                    r1 = L
                elif lock == "stable":
                    L = clamp(r2)
                    rem = max(0.0, 100.0 - L)
                    if changed == "runin":
                        r1 = clamp(r1, 0.0, rem)
                        r3 = rem - r1
                    elif changed == "severe":
                        r3 = clamp(r3, 0.0, rem)
                        r1 = rem - r3
                    else:
                        tot = r1 + r3
                        if tot <= 1e-9:
                            r1 = rem * 0.5
                            r3 = rem - r1
                        else:
                            r1 = rem * (r1 / tot)
                            r3 = rem - r1
                    r2 = L
                else:  # lock == severe
                    L = clamp(r3)
                    rem = max(0.0, 100.0 - L)
                    if changed == "runin":
                        r1 = clamp(r1, 0.0, rem)
                        r2 = rem - r1
                    elif changed == "stable":
                        r2 = clamp(r2, 0.0, rem)
                        r1 = rem - r2
                    else:
                        tot = r1 + r2
                        if tot <= 1e-9:
                            r1 = rem * 0.5
                            r2 = rem - r1
                        else:
                            r1 = rem * (r1 / tot)
                            r2 = rem - r1
                    r3 = L

                # 最后做一次数值保护，避免浮点误差导致 Σ!=100
                s = r1 + r2 + r3
                if abs(s - 100.0) > 1e-6:
                    # 优先把误差丢到未锁定且不是当前 changed 的那一段（如果存在）
                    if lock != "runin" and changed != "runin":
                        r1 = max(0.0, 100.0 - r2 - r3)
                    elif lock != "stable" and changed != "stable":
                        r2 = max(0.0, 100.0 - r1 - r3)
                    else:
                        r3 = max(0.0, 100.0 - r1 - r2)

            self.phase_runin_pct.set(r1)
            self.phase_stable_pct.set(r2)
            self.phase_severe_pct.set(r3)
            self._update_ratio_labels()
        finally:
            self._ratio_lock = False

    
    def _on_ratio_entry(self, which: str):
        """从输入框读取百分比并应用（回车或失焦触发）"""
        if getattr(self, "_ratio_lock", False):
            return
        try:
            if which == "runin":
                v = float(self._lbl_runin.get().strip())
                self.phase_runin_pct.set(v)
            elif which == "stable":
                v = float(self._lbl_stable.get().strip())
                self.phase_stable_pct.set(v)
            else:
                v = float(self._lbl_severe.get().strip())
                self.phase_severe_pct.set(v)
        except Exception:
            # 输入非法则回填当前值
            self._update_ratio_labels()
            return
        self._on_ratio_change(which)

    def _on_lock_change(self):
        """锁定一个阶段：禁用对应滑块与输入框"""
        lock = self.locked_phase.get() if hasattr(self, "locked_phase") else "none"

        def set_widget_state(w, enabled: bool):
            if w is None:
                return
            try:
                w.state(["!disabled"] if enabled else ["disabled"])
            except Exception:
                try:
                    w.configure(state=("normal" if enabled else "disabled"))
                except Exception:
                    pass

        # 默认都可编辑
        set_widget_state(getattr(self, "_scale_runin", None), lock != "runin")
        set_widget_state(getattr(self, "_entry_runin", None), lock != "runin")

        set_widget_state(getattr(self, "_scale_stable", None), lock != "stable")
        set_widget_state(getattr(self, "_entry_stable", None), lock != "stable")

        set_widget_state(getattr(self, "_scale_severe", None), lock != "severe")
        set_widget_state(getattr(self, "_entry_severe", None), lock != "severe")

        # 锁定切换后，强制归一化一次（避免 Σ != 100）
        self._on_ratio_change("runin")
    def _build_run_panel(self, parent):
        box = ttk.LabelFrame(parent, text="导出与输出", padding=10)
        box.pack(fill="x")

        ttk.Label(box, text="输出目录").grid(row=0, column=0, sticky="w")
        ttk.Entry(box, textvariable=self.out_dir, width=34).grid(row=1, column=0, sticky="we", pady=(4, 6))
        ttk.Button(box, text="选择…", command=self._choose_dir).grid(row=1, column=1, padx=(6, 0))

        ttk.Label(box, text="随机种子 seed").grid(row=2, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(box, textvariable=self.seed, width=12).grid(row=3, column=0, sticky="w", pady=(4, 6))

        # 图片语言选择
        lang_box = ttk.Frame(box)
        lang_box.grid(row=4, column=0, columnspan=2, sticky="we", pady=(6, 2))
        ttk.Label(lang_box, text="图片语言").pack(side="left")
        ttk.Radiobutton(lang_box, text="中文", variable=self.plot_lang, value="zh").pack(side="left", padx=(10, 0))
        ttk.Radiobutton(lang_box, text="英文", variable=self.plot_lang, value="en").pack(side="left", padx=(10, 0))

        # 两个按钮：导出 xlsx / 导出 图片
        self.btn_xlsx = ttk.Button(box, text="导出 xlsx", command=self._start_export_xlsx)
        self.btn_xlsx.grid(row=5, column=0, sticky="we", pady=(8, 4))
        self.btn_plots = ttk.Button(box, text="导出 图片", command=self._start_export_plots)
        self.btn_plots.grid(row=5, column=1, sticky="we", pady=(8, 4), padx=(6, 0))

        # 进度条（确定型） + 百分比
        self.pb = ttk.Progressbar(box, mode="determinate", maximum=100.0)
        self.pb.grid(row=6, column=0, sticky="we", pady=(6, 3))
        self.pb_pct = tk.StringVar(value="0%")
        ttk.Label(box, textvariable=self.pb_pct).grid(row=6, column=1, padx=(6, 0), sticky="w")

        self.log = tk.Text(parent, height=26, width=54)
        self.log.pack(fill="both", expand=True, pady=(10, 0))
        self.log.insert("end", "日志输出：\n")
        self.log.configure(state="disabled")

        tips = ("提示：\n"
                "1) 建议安装 SciPy：pip install scipy（陷波/低通更快更标准）\n"
                "2) duration 很长（例如10小时@50Hz）时，xlsx 会很大，可把 stride 设为 2~10。\n"
                "3) 若中文图仍乱码：请安装中文字体（微软雅黑/黑体/Noto Sans CJK/文泉驿等）。\n")
        ttk.Label(parent, text=tips).pack(anchor="w", pady=(10, 0))

    def _choose_dir(self):
        d = filedialog.askdirectory(title="选择输出目录")
        if d:
            self.out_dir.set(d)

    def _read_cfg(self):
        for k, v in self._vars.items():
            s = v.get().strip()
            try:
                typ = self._types.get(k, float)
                val = typ(float(s))
                if k == "duration_h":
                    # GUI 以小时输入，内部仍使用秒
                    setattr(self.cfg, "duration_s", float(val) * 3600.0)
                else:
                    setattr(self.cfg, k, val)
            except Exception:
                raise ValueError(f"字段 {k} 输入不合法：{s}")
        # 阶段比例（GUI滑块，单位%）
        try:
            self.cfg.phase_runin_ratio = float(self.phase_runin_pct.get()) / 100.0
            self.cfg.phase_stable_ratio = float(self.phase_stable_pct.get()) / 100.0
            self.cfg.phase_severe_ratio = float(self.phase_severe_pct.get()) / 100.0
        except Exception:
            pass
        self.cfg.validate()

    def _make_cache_key(self) -> str:
        items = [
            ("theta", self.cfg.theta_deg),
            ("tset", self.cfg.t_set_N),
            ("fs", self.cfg.fs_Hz),
            ("dur_h", self.cfg.duration_s / 3600.0),
            ("rpm", self.cfg.rpm),
            ("m", self.cfg.mech_harmonic),

            ("noise_o", self.cfg.noise_rms_open),
            ("noise_c", self.cfg.noise_rms_closed),
            ("mechAo", self.cfg.mech_amp_open),
            ("mechAc", self.cfg.mech_amp_closed),
            ("driftAo", self.cfg.drift_amp_open),
            ("driftAc", self.cfg.drift_amp_closed),
            ("driftf", self.cfg.drift_freq_hz),
            ("sensor", self.cfg.sensor_rms),

            ("hws", self.cfg.hampel_win_s),
            ("hns", self.cfg.hampel_nsig),
            ("fc", self.cfg.lowpass_fc_hz),

            ("Wss", self.cfg.stable_win_s),
            ("sig", self.cfg.stable_sigma_max),
            ("slope", self.cfg.stable_slope_max),
            ("qmin", self.cfg.stable_valid_min),
            ("delta", self.cfg.fail_delta),
            ("hold", self.cfg.fail_hold_s),

            ("stride", self.cfg.export_stride),
            ("pmax", self.cfg.plot_max_points),

            # 阶段比例/μ范围（GUI新增）
            ("phase_r1", float(getattr(self.cfg, "phase_runin_ratio", 0.0))),
            ("phase_r2", float(getattr(self.cfg, "phase_stable_ratio", 0.0))),
            ("phase_r3", float(getattr(self.cfg, "phase_severe_ratio", 0.0))),
            ("mu_r_min", float(getattr(self.cfg, "mu_runin_min", 0.0))),
            ("mu_r_max", float(getattr(self.cfg, "mu_runin_max", 0.0))),
            ("mu_s_min", float(getattr(self.cfg, "mu_stable_min", 0.0))),
            ("mu_s_max", float(getattr(self.cfg, "mu_stable_max", 0.0))),
            ("mu_a_min", float(getattr(self.cfg, "mu_severe_min", 0.0))),
            ("mu_a_max", float(getattr(self.cfg, "mu_severe_max", 0.0))),

            ("seed", int(self.seed.get())),
        ]
        return "|".join([f"{k}={v}" for k, v in items])

    def _ensure_simulated(self, progress_cb):
        key = self._make_cache_key()
        if self._cache_res is not None and self._cache_key == key:
            progress_cb(10.0, "参数未变，使用缓存的仿真结果（跳过重算）")
            return self._cache_res

        progress_cb(1.0, "开始生成仿真数据（用于导出）...")
        res = simulate(self.cfg, seed=int(self.seed.get()), progress_cb=progress_cb)
        self._cache_key = key
        self._cache_res = res
        return res

    def _start_export_xlsx(self):
        self._start_task(task="xlsx")

    def _start_export_plots(self):
        self._start_task(task="plots")

    def _start_task(self, task: str):
        if self.worker and self.worker.is_alive():
            messagebox.showinfo("提示", "正在运行中，请等待完成。")
            return
        try:
            self._read_cfg()
        except Exception as e:
            messagebox.showerror("输入错误", str(e))
            return

        out_dir = self.out_dir.get().strip()
        if not out_dir:
            messagebox.showerror("输入错误", "请设置输出目录。")
            return

        self.btn_xlsx.configure(state="disabled")
        self.btn_plots.configure(state="disabled")
        self.pb["value"] = 0.0
        self.pb_pct.set("0%")
        self.status.set("运行中…")
        self._log(f"开始任务：导出 {task}\n")

        def progress_cb(pct, msg):
            self.msg_q.put(("progress", float(pct), str(msg)))

        def work():
            try:
                res = self._ensure_simulated(progress_cb)
                outputs = {}
                if task == "xlsx":
                    outputs["xlsx"] = export_xlsx(res, out_dir=out_dir, progress_cb=progress_cb)
                elif task == "plots":
                    outputs.update(export_plots(res, out_dir=out_dir, lang=self.plot_lang.get(), progress_cb=progress_cb))
                export_summary(res, out_dir=out_dir, extra={"outputs": outputs})
                self.msg_q.put(("done", task))
            except Exception as e:
                self.msg_q.put(("err", str(e)))

        self.worker = threading.Thread(target=work, daemon=True)
        self.worker.start()

    def _poll_msgs(self):
        try:
            while True:
                item = self.msg_q.get_nowait()
                typ = item[0]
                if typ == "progress":
                    _, pct, msg = item
                    pct = max(0.0, min(100.0, float(pct)))
                    self.pb["value"] = pct
                    self.pb_pct.set(f"{pct:.0f}%")
                    self.status.set(f"运行中… {pct:.0f}%")
                    if msg:
                        self._log(f"[{pct:6.2f}%] {msg}\n")
                elif typ == "done":
                    _, task = item
                    self.btn_xlsx.configure(state="normal")
                    self.btn_plots.configure(state="normal")
                    self.pb["value"] = 100.0
                    self.pb_pct.set("100%")
                    self.status.set("完成")
                    self._log(f"完成：{task}\n")
                    messagebox.showinfo("完成", f"已完成导出：{task}\n请到输出目录查看。")
                elif typ == "err":
                    self.btn_xlsx.configure(state="normal")
                    self.btn_plots.configure(state="normal")
                    self.status.set("出错")
                    self._log("错误：" + item[1] + "\n")
                    messagebox.showerror("运行出错", item[1])
        except queue.Empty:
            pass
        self.after(120, self._poll_msgs)

    def _log(self, s: str):
        self.log.configure(state="normal")
        self.log.insert("end", s)
        self.log.see("end")
        self.log.configure(state="disabled")


if __name__ == "__main__":
    App().mainloop()
