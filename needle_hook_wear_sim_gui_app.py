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
import configparser
from pathlib import Path
import threading
import queue
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from needle_hook_wear_simulator_gui import (
    SimConfig,
    simulate,
    export_data,
    export_plots,
    export_summary,
    setup_plot_font,
    HAVE_SCIPY,
    EXPORT_COLUMN_SPECS,
    DEFAULT_EXPORT_COLUMNS,
    EXPORT_TABLE_SPECS,
    DEFAULT_EXPORT_TABLES,
)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("针钩磨损检测全过程仿真（Capstan + 扰动 + 滤波 + 判据）")
        # 设置窗口左上角图标（Windows）
        try:
            base_dir = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
            icon_path = base_dir / "app.ico"
            if icon_path.exists():
                self.iconbitmap(str(icon_path))
        except Exception:
            pass
        self.geometry("1080x760")
        self.minsize(980, 700)

        self.msg_q = queue.Queue()
        self.worker = None
        self._init_ini_paths()

        self.cfg = SimConfig()  # 默认包角100°，默认rpm=300
        self.out_dir = tk.StringVar(value=os.path.abspath("sim_out"))
        self.seed = tk.IntVar(value=7)

        self.plot_lang = tk.StringVar(value="zh")  # zh / en
        self.data_export_format = tk.StringVar(value="xlsx")
        self.export_column_vars = {
            key: tk.BooleanVar(value=(key in DEFAULT_EXPORT_COLUMNS))
            for key, _ in EXPORT_COLUMN_SPECS
        }
        self.export_table_vars = {
            key: tk.BooleanVar(value=(key in DEFAULT_EXPORT_TABLES))
            for key, _ in EXPORT_TABLE_SPECS
        }
        self.invalid_ratio_pct = tk.StringVar(value=f"{float(getattr(self.cfg, 'invalid_ratio', 0.0005)) * 100.0:.6g}")
        self.phase_control_mode = tk.StringVar(value=str(getattr(self.cfg, "phase_control_mode", "ratio")))
        self.phase_runin_tlife_pct = tk.DoubleVar(value=float(getattr(self.cfg, "runin_ratio_by_tlife", self.cfg.phase_runin_ratio)) * 100.0)
        self._lbl_runin_tlife = tk.StringVar()
        self.target_tlife_s = tk.StringVar(value=f"{float(getattr(self.cfg, 'target_tlife_s', self.cfg.duration_s * (self.cfg.phase_runin_ratio + self.cfg.phase_stable_ratio))):.6f}")
        self.phase_tlife_info = tk.StringVar()

        # 开环-闭环对比输出范围（单位：s），(0,-1)=全部
        self.compare_t_start_s = tk.StringVar(value=str(getattr(self.cfg, "compare_t_start_s", 0.0)))
        self.compare_t_end_s = tk.StringVar(value=str(getattr(self.cfg, "compare_t_end_s", -1.0)))

        # 缓存：避免连续导出重复计算
        self._cache_key = None
        self._cache_res = None

        top = ttk.Frame(self, padding=10)
        top.pack(fill="x")
        info = ("说明：fs 与采样时间仅用于生成时间轴（输出点数）。\n"
                "机械主频：仅由 rpm 换算 f_mech=rpm/60*m。\n"
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

        # 读取/生成 defaultData.ini（UTF-8，含中文注释）
        self._load_or_create_default_ini()
        self.protocol("WM_DELETE_WINDOW", self.destroy)

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
        tab_phase_host, tab_phase = self._create_scrollable_tab(nb, padding=10)
        tab_filter = ttk.Frame(nb, padding=10)
        tab_judge = ttk.Frame(nb, padding=10)
        tab_export = ttk.Frame(nb, padding=10)
        nb.add(tab_core, text="核心输入")
        nb.add(tab_dist, text="扰动参数")
        nb.add(tab_phase_host, text="阶段比例/μ范围")
        nb.add(tab_filter, text="滤波参数")
        nb.add(tab_judge, text="基线/阈值/寿命")
        nb.add(tab_export, text="导出/绘图")

        self._vars = {}
        self._types = {}

        # 核心输入
        self._add_entry(tab_core, "包角 θ (deg)", "theta_deg")
        self._add_entry(tab_core, "平均张力设定 T_set (N)", "t_set_N")
        self._add_entry(tab_core, "采样率 fs (Hz) 仅生成时间轴", "fs_Hz")
        self._add_entry(tab_core, "采样时间 duration (h) 仅生成时间轴", "duration_h")


        # 阶段控制方式
        tab_phase.columnconfigure(1, weight=1)
        ttk.Label(tab_phase, text="阶段控制方式").grid(row=0, column=0, sticky="w", pady=(0, 4))
        mode_box = ttk.Frame(tab_phase)
        mode_box.grid(row=0, column=1, columnspan=3, sticky="w", pady=(0, 4))
        self._phase_mode_btn_ratio = ttk.Radiobutton(mode_box, text="阶段时间比例", variable=self.phase_control_mode, value="ratio", command=self._update_phase_mode_state)
        self._phase_mode_btn_ratio.pack(side="left", padx=(0, 10))
        self._phase_mode_btn_tlife = ttk.Radiobutton(mode_box, text="磨合比例 + tlife", variable=self.phase_control_mode, value="runin_tlife", command=self._update_phase_mode_state)
        self._phase_mode_btn_tlife.pack(side="left")
        ttk.Label(tab_phase, text="方式1直接控制三段时间比例；方式2控制磨合段比例和 tlife，并自动推导稳定/加速比例。").grid(row=1, column=0, columnspan=4, sticky="w", pady=(0, 8))

        self._ratio_lock = False
        self.phase_runin_pct = tk.DoubleVar(value=float(self.cfg.phase_runin_ratio) * 100.0)
        self.phase_stable_pct = tk.DoubleVar(value=float(self.cfg.phase_stable_ratio) * 100.0)
        self.phase_severe_pct = tk.DoubleVar(value=float(self.cfg.phase_severe_ratio) * 100.0)
        self.locked_phase = tk.StringVar(value="none")  # none/runin/stable/severe
        self._lbl_runin = tk.StringVar()
        self._lbl_stable = tk.StringVar()
        self._lbl_severe = tk.StringVar()
        self.phase_sum_label = tk.StringVar()

        ttk.Label(tab_phase, text="阶段比例（%）").grid(row=2, column=0, columnspan=4, sticky="w", pady=(0, 4))

        ttk.Label(tab_phase, text="磨合阶段").grid(row=3, column=0, sticky="w", pady=4)
        self._scale_runin = ttk.Scale(tab_phase, from_=0.0, to=100.0, variable=self.phase_runin_pct,
                                      command=lambda _v: self._on_ratio_change("runin"))
        self._scale_runin.grid(row=3, column=1, sticky="we", pady=4, padx=(0, 8))
        self._entry_runin = ttk.Entry(tab_phase, textvariable=self._lbl_runin, width=8)
        self._entry_runin.grid(row=3, column=2, sticky="w")
        ttk.Label(tab_phase, text="%").grid(row=3, column=3, sticky="w", padx=(4, 0))
        self._entry_runin.bind("<Return>", lambda _e: self._on_ratio_entry("runin"))
        self._entry_runin.bind("<FocusOut>", lambda _e: self._on_ratio_entry("runin"))

        ttk.Label(tab_phase, text="稳定磨损阶段").grid(row=4, column=0, sticky="w", pady=4)
        self._scale_stable = ttk.Scale(tab_phase, from_=0.0, to=100.0, variable=self.phase_stable_pct,
                                       command=lambda _v: self._on_ratio_change("stable"))
        self._scale_stable.grid(row=4, column=1, sticky="we", pady=4, padx=(0, 8))
        self._entry_stable = ttk.Entry(tab_phase, textvariable=self._lbl_stable, width=8)
        self._entry_stable.grid(row=4, column=2, sticky="w")
        ttk.Label(tab_phase, text="%").grid(row=4, column=3, sticky="w", padx=(4, 0))
        self._entry_stable.bind("<Return>", lambda _e: self._on_ratio_entry("stable"))
        self._entry_stable.bind("<FocusOut>", lambda _e: self._on_ratio_entry("stable"))

        ttk.Label(tab_phase, text="加速磨损阶段").grid(row=5, column=0, sticky="w", pady=4)
        self._scale_severe = ttk.Scale(tab_phase, from_=0.0, to=100.0, variable=self.phase_severe_pct,
                                       command=lambda _v: self._on_ratio_change("severe"))
        self._scale_severe.grid(row=5, column=1, sticky="we", pady=4, padx=(0, 8))
        self._entry_severe = ttk.Entry(tab_phase, textvariable=self._lbl_severe, width=8)
        self._entry_severe.grid(row=5, column=2, sticky="w")
        ttk.Label(tab_phase, text="%").grid(row=5, column=3, sticky="w", padx=(4, 0))
        self._entry_severe.bind("<Return>", lambda _e: self._on_ratio_entry("severe"))
        self._entry_severe.bind("<FocusOut>", lambda _e: self._on_ratio_entry("severe"))

        ttk.Label(tab_phase, textvariable=self.phase_sum_label).grid(row=6, column=0, columnspan=4, sticky="w", pady=(4, 6))

        ttk.Label(tab_phase, text="锁定一个阶段").grid(row=7, column=0, sticky="w", pady=4)
        lock_box = ttk.Frame(tab_phase)
        lock_box.grid(row=7, column=1, columnspan=3, sticky="w", pady=4)
        self._lock_btn_none = ttk.Radiobutton(lock_box, text="不锁定", variable=self.locked_phase, value="none", command=self._on_lock_change)
        self._lock_btn_none.pack(side="left", padx=(0, 10))
        self._lock_btn_runin = ttk.Radiobutton(lock_box, text="锁定磨合", variable=self.locked_phase, value="runin", command=self._on_lock_change)
        self._lock_btn_runin.pack(side="left", padx=(0, 10))
        self._lock_btn_stable = ttk.Radiobutton(lock_box, text="锁定稳定", variable=self.locked_phase, value="stable", command=self._on_lock_change)
        self._lock_btn_stable.pack(side="left", padx=(0, 10))
        self._lock_btn_severe = ttk.Radiobutton(lock_box, text="锁定加速", variable=self.locked_phase, value="severe", command=self._on_lock_change)
        self._lock_btn_severe.pack(side="left", padx=(0, 10))
        self._lock_buttons = [self._lock_btn_none, self._lock_btn_runin, self._lock_btn_stable, self._lock_btn_severe]

        alt_box = ttk.LabelFrame(tab_phase, text="磨合比例 + tlife 控制", padding=8)
        alt_box.grid(row=8, column=0, columnspan=4, sticky="we", pady=(6, 10))
        alt_box.columnconfigure(1, weight=1)
        ttk.Label(alt_box, text="磨合段比例").grid(row=0, column=0, sticky="w", pady=4)
        self._scale_runin_tlife = ttk.Scale(alt_box, from_=0.0, to=100.0, variable=self.phase_runin_tlife_pct,
                                            command=lambda _v: self._on_runin_tlife_change())
        self._scale_runin_tlife.grid(row=0, column=1, sticky="we", pady=4, padx=(0, 8))
        self._entry_runin_tlife = ttk.Entry(alt_box, textvariable=self._lbl_runin_tlife, width=8)
        self._entry_runin_tlife.grid(row=0, column=2, sticky="w")
        ttk.Label(alt_box, text="%").grid(row=0, column=3, sticky="w", padx=(4, 0))
        self._entry_runin_tlife.bind("<Return>", lambda _e: self._on_runin_tlife_entry())
        self._entry_runin_tlife.bind("<FocusOut>", lambda _e: self._on_runin_tlife_entry())

        ttk.Label(alt_box, text="目标 tlife (s)").grid(row=1, column=0, sticky="w", pady=4)
        self._entry_target_tlife = ttk.Entry(alt_box, textvariable=self.target_tlife_s, width=12)
        self._entry_target_tlife.grid(row=1, column=1, sticky="w", pady=4)
        self._entry_target_tlife.bind("<Return>", lambda _e: self._update_runin_tlife_labels())
        self._entry_target_tlife.bind("<FocusOut>", lambda _e: self._update_runin_tlife_labels())
        ttk.Label(alt_box, textvariable=self.phase_tlife_info).grid(row=2, column=0, columnspan=4, sticky="w", pady=(4, 0))

        ttk.Label(tab_phase, text="三阶段摩擦系数范围（min/max；若 min==max 即固定值）").grid(row=9, column=0, columnspan=4, sticky="w", pady=(0, 8))
        self._add_entry(tab_phase, "磨合段 μ范围：min", "mu_runin_min", row=10)
        self._add_entry(tab_phase, "磨合段 μ范围：max", "mu_runin_max", row=11)
        self._add_entry(tab_phase, "稳定段 μ范围：min", "mu_stable_min", row=12)
        self._add_entry(tab_phase, "稳定段 μ范围：max", "mu_stable_max", row=13)
        self._add_entry(tab_phase, "加速段 μ范围：min", "mu_severe_min", row=14)
        self._add_entry(tab_phase, "加速段 μ范围：max", "mu_severe_max", row=15)
        self._add_entry(tab_phase, "磨合→稳定过渡系数 k_rs（越大越快）", "trans_runin2stable_k", row=16)
        self._add_entry(tab_phase, "稳定→加速过渡系数 k_sa（越大越快）", "trans_stable2severe_k", row=17)

        self._update_ratio_labels()
        self._update_runin_tlife_labels()
        if "duration_h" in self._vars:
            self._vars["duration_h"].trace_add("write", lambda *_: self._update_runin_tlife_labels())
        self._update_phase_mode_state()

        # 扰动：转速范围输入主频（删除直接输入主频）
        ttk.Label(tab_dist, text="机械周期扰动主频：f_mech = rpm/60*m（仅此方式输入）").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))
        self._add_range_entry(tab_dist, "转速范围 rpm (min~max)", "rpm_min", "rpm_max", row=1)
        self._add_entry(tab_dist, "倍频 m（1=一次转频）", "mech_harmonic", row=2, is_int=True)

        self._add_range_entry(tab_dist, "开环周期幅值 A_mech_open (N)", "mech_amp_open_min", "mech_amp_open_max", row=3)
        self._add_range_entry(tab_dist, "闭环周期幅值 A_mech_closed (N)", "mech_amp_closed_min", "mech_amp_closed_max", row=4)
        self._add_range_entry(tab_dist, "开环噪声 RMS_open (N)", "noise_rms_open_min", "noise_rms_open_max", row=5)
        self._add_range_entry(tab_dist, "闭环噪声 RMS_closed (N)", "noise_rms_closed_min", "noise_rms_closed_max", row=6)
        self._add_range_entry(tab_dist, "漂移频率 f_drift (Hz，准周期)", "drift_freq_hz_min", "drift_freq_hz_max", row=7)
        self._add_range_entry(tab_dist, "开环漂移幅值 A_drift_open (N)", "drift_amp_open_min", "drift_amp_open_max", row=8)
        self._add_range_entry(tab_dist, "闭环漂移幅值 A_drift_closed (N)", "drift_amp_closed_min", "drift_amp_closed_max", row=9)
        self._add_range_entry(tab_dist, "传感器噪声 RMS_sensor (N)", "sensor_rms_min", "sensor_rms_max", row=10)

        self._add_entry(tab_dist, "转速慢变时间常数 τ_rpm (s，<=0关闭)", "tau_rpm_s", row=11)
        self._add_entry(tab_dist, "机械扰动慢变时间常数 τ_mech (s，<=0关闭)", "tau_mech_s", row=12)
        self._add_entry(tab_dist, "噪声RMS慢变时间常数 τ_noise (s，<=0关闭)", "tau_noise_s", row=13)
        self._add_entry(tab_dist, "传感器RMS慢变时间常数 τ_sensor (s，<=0关闭)", "tau_sensor_s", row=14)
        self._add_entry(tab_dist, "漂移幅值慢变时间常数 τ_driftA (s，<=0关闭)", "tau_drift_amp_s", row=15)
        self._add_entry(tab_dist, "漂移频率慢变时间常数 τ_driftf (s，<=0关闭)", "tau_drift_freq_s", row=16)


        # 显示 f_mech 与 Notch Q（自动，派生信息，只读；放在该页最下方，避免与输入框同一行）
        self.f_mech_var = tk.StringVar(value="f_mech = — Hz，Notch Q≈—")
        info_row = tab_dist.grid_size()[1]  # 追加到末尾（下一空行）
        ttk.Label(tab_dist, textvariable=self.f_mech_var).grid(row=info_row, column=0, columnspan=2, sticky="w", pady=(10, 0))
        self._vars["rpm_min"].trace_add("write", lambda *_: self._update_f_mech_label())
        self._vars["rpm_max"].trace_add("write", lambda *_: self._update_f_mech_label())
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
        self._add_entry(tab_judge, "最短连续稳定段 Whold (s)", "stable_hold_s", row=2)
        self._add_entry(tab_judge, "稳定标准差 σ_max", "stable_sigma_max", row=3)
        self._add_entry(tab_judge, "稳定总漂移阈值 Δμ_max", "stable_slope_max", row=4)
        self._add_entry(tab_judge, "稳定有效比例 q_min", "stable_valid_min", row=5)
        self._add_entry(tab_judge, "失效阈值增量 δ（μth=μss*(1+δ)）", "fail_delta", row=6)
        self._add_entry(tab_judge, "持续超限 Wpersist (s)", "fail_hold_s", row=7)
        ttk.Label(tab_judge, text="invalid 比例（%，quality flag=0）").grid(row=8, column=0, sticky="w", padx=(0, 10), pady=4)
        ttk.Entry(tab_judge, textvariable=self.invalid_ratio_pct, width=24).grid(row=8, column=1, sticky="w", pady=4)

        # 导出/绘图
        ttk.Label(tab_export, text="duration 很长时 xlsx 会很大；可用 stride 抽样导出。绘图会自动降采样。").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))
        self._add_entry(tab_export, "导出步长 stride（1=全量）", "export_stride", row=1, is_int=True)
        self._add_entry(tab_export, "绘图最大点数", "plot_max_points", row=2, is_int=True)

        tab_export.columnconfigure(1, weight=1)
        fmt_box = ttk.Frame(tab_export)
        fmt_box.grid(row=3, column=0, columnspan=2, sticky="w", pady=(4, 4))
        ttk.Label(fmt_box, text="数据导出格式").pack(side="left")
        ttk.Radiobutton(fmt_box, text="xlsx", variable=self.data_export_format, value="xlsx").pack(side="left", padx=(10, 0))
        ttk.Radiobutton(fmt_box, text="SQLite (.db)", variable=self.data_export_format, value="db").pack(side="left", padx=(10, 0))
        self._build_export_table_checks(tab_export, row=4)
        self._build_export_column_checks(tab_export, row=5)

    def _create_scrollable_tab(self, notebook, padding=10):
        host = ttk.Frame(notebook)
        host.columnconfigure(0, weight=1)
        host.rowconfigure(0, weight=1)

        canvas = tk.Canvas(host, highlightthickness=0, borderwidth=0)
        ysb = ttk.Scrollbar(host, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=ysb.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        ysb.grid(row=0, column=1, sticky="ns")

        body = ttk.Frame(canvas, padding=padding)
        window_id = canvas.create_window((0, 0), window=body, anchor="nw")

        def _sync_scrollregion(_event=None):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _fit_width(event):
            canvas.itemconfigure(window_id, width=event.width)

        def _on_mousewheel(event):
            if event.delta:
                step = -int(event.delta / 120)
            elif getattr(event, "num", None) == 4:
                step = -1
            elif getattr(event, "num", None) == 5:
                step = 1
            else:
                step = 0
            if step:
                canvas.yview_scroll(step, "units")
                return "break"
            return None

        def _bind_mousewheel(_event=None):
            canvas.bind_all("<MouseWheel>", _on_mousewheel)
            canvas.bind_all("<Button-4>", _on_mousewheel)
            canvas.bind_all("<Button-5>", _on_mousewheel)

        def _unbind_mousewheel(_event=None):
            canvas.unbind_all("<MouseWheel>")
            canvas.unbind_all("<Button-4>")
            canvas.unbind_all("<Button-5>")

        body.bind("<Configure>", _sync_scrollregion)
        canvas.bind("<Configure>", _fit_width)
        for widget in (canvas, body):
            widget.bind("<Enter>", _bind_mousewheel, add="+")
            widget.bind("<Leave>", _unbind_mousewheel, add="+")

        return host, body

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

    def _add_range_entry(self, parent, label, field_min, field_max, row=None):
        """添加范围输入：两个输入框（min/max）。"""
        if row is None:
            row = parent.grid_size()[1]
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=(0, 10), pady=4)
        frm = ttk.Frame(parent)
        frm.grid(row=row, column=1, sticky="w", pady=4)
        vmin = tk.StringVar(value=str(getattr(self.cfg, field_min)))
        ttk.Entry(frm, textvariable=vmin, width=10).grid(row=0, column=0, sticky="w")
        ttk.Label(frm, text="~").grid(row=0, column=1, padx=6)
        vmax = tk.StringVar(value=str(getattr(self.cfg, field_max)))
        ttk.Entry(frm, textvariable=vmax, width=10).grid(row=0, column=2, sticky="w")
        self._vars[field_min] = vmin
        self._vars[field_max] = vmax
        self._types[field_min] = float
        self._types[field_max] = float

    def _get_selected_export_columns(self):
        return [
            key
            for key, _ in EXPORT_COLUMN_SPECS
            if self.export_column_vars[key].get()
        ]

    def _get_selected_export_tables(self):
        return [
            key
            for key, _ in EXPORT_TABLE_SPECS
            if self.export_table_vars[key].get()
        ]

    def _export_columns_to_ini_value(self):
        return ",".join(self._get_selected_export_columns())

    def _export_tables_to_ini_value(self):
        return ",".join(self._get_selected_export_tables())

    def _set_export_columns_from_csv(self, raw_value):
        raw = str(raw_value or "").strip()
        tokens = {part.strip() for part in raw.split(",") if part.strip()}
        selected = [key for key, _ in EXPORT_COLUMN_SPECS if key in tokens]
        if not selected:
            selected = list(DEFAULT_EXPORT_COLUMNS)
        selected_set = set(selected)
        for key, _ in EXPORT_COLUMN_SPECS:
            self.export_column_vars[key].set(key in selected_set)

    def _set_export_tables_from_csv(self, raw_value):
        raw = str(raw_value or "").strip()
        tokens = {part.strip() for part in raw.split(",") if part.strip()}
        selected = [key for key, _ in EXPORT_TABLE_SPECS if key in tokens]
        if not selected:
            selected = list(DEFAULT_EXPORT_TABLES)
        selected_set = set(selected)
        for key, _ in EXPORT_TABLE_SPECS:
            self.export_table_vars[key].set(key in selected_set)

    def _build_export_table_checks(self, parent, row):
        box = ttk.LabelFrame(parent, text="数据导出表", padding=8)
        box.grid(row=row, column=0, columnspan=2, sticky="we", pady=(6, 0))
        for col in range(2):
            box.columnconfigure(col, weight=1)

        for idx, (key, label) in enumerate(EXPORT_TABLE_SPECS):
            r = idx // 2
            c = idx % 2
            ttk.Checkbutton(
                box,
                text=label,
                variable=self.export_table_vars[key],
            ).grid(row=r, column=c, sticky="w", padx=(0, 12), pady=2)

    def _build_export_column_checks(self, parent, row):
        box = ttk.LabelFrame(parent, text="数据导出列", padding=8)
        box.grid(row=row, column=0, columnspan=2, sticky="we", pady=(6, 0))
        for col in range(3):
            box.columnconfigure(col, weight=1)

        for idx, (key, label) in enumerate(EXPORT_COLUMN_SPECS):
            r = idx // 3
            c = idx % 3
            ttk.Checkbutton(
                box,
                text=label,
                variable=self.export_column_vars[key],
            ).grid(row=r, column=c, sticky="w", padx=(0, 12), pady=2)

    def _set_widget_state(self, w, enabled: bool):
        if w is None:
            return
        try:
            w.state(["!disabled"] if enabled else ["disabled"])
        except Exception:
            try:
                w.configure(state=("normal" if enabled else "disabled"))
            except Exception:
                pass

    def _update_runin_tlife_labels(self):
        try:
            runin_pct = max(0.0, min(100.0, float(self.phase_runin_tlife_pct.get())))
        except Exception:
            runin_pct = 0.0
        self.phase_runin_tlife_pct.set(runin_pct)
        self._lbl_runin_tlife.set(f"{runin_pct:.2f}")

        try:
            duration_h = float(self._vars["duration_h"].get().strip())
            duration_s = max(0.0, duration_h * 3600.0)
        except Exception:
            duration_s = float(getattr(self.cfg, "duration_s", 0.0))

        try:
            tlife_s = float(self.target_tlife_s.get().strip())
        except Exception:
            tlife_s = duration_s * max(0.0, min(1.0, runin_pct / 100.0))

        if duration_s <= 0.0:
            self.phase_tlife_info.set("待输入有效 duration 和 tlife 后，可自动推导稳定/加速比例。")
            return

        fs = max(1e-9, float(getattr(self.cfg, "fs_Hz", 1.0)))
        min_tlife = min(duration_s, duration_s * (runin_pct / 100.0) + 1.0 / fs)
        tlife_s = max(min_tlife, min(duration_s, tlife_s))
        stable_pct = max(0.0, (tlife_s / duration_s) * 100.0 - runin_pct)
        severe_pct = max(0.0, 100.0 - runin_pct - stable_pct)
        self.phase_tlife_info.set(
            f"按当前 duration 推导：稳定≈{stable_pct:.2f}% / 加速≈{severe_pct:.2f}%；tlife={tlife_s:.2f}s"
        )

    def _on_runin_tlife_change(self):
        self._update_runin_tlife_labels()

    def _on_runin_tlife_entry(self):
        try:
            self.phase_runin_tlife_pct.set(float(self._lbl_runin_tlife.get().strip()))
        except Exception:
            pass
        self._update_runin_tlife_labels()

    def _update_phase_mode_state(self):
        ratio_mode = (self.phase_control_mode.get() == "ratio")
        lock = self.locked_phase.get() if hasattr(self, "locked_phase") else "none"

        self._set_widget_state(getattr(self, "_phase_mode_btn_ratio", None), True)
        self._set_widget_state(getattr(self, "_phase_mode_btn_tlife", None), True)

        for btn in getattr(self, "_lock_buttons", []):
            self._set_widget_state(btn, ratio_mode)

        if ratio_mode:
            self._on_lock_change()
        else:
            for w in [
                getattr(self, "_scale_runin", None),
                getattr(self, "_entry_runin", None),
                getattr(self, "_scale_stable", None),
                getattr(self, "_entry_stable", None),
                getattr(self, "_scale_severe", None),
                getattr(self, "_entry_severe", None),
            ]:
                self._set_widget_state(w, False)

        self._set_widget_state(getattr(self, "_scale_runin_tlife", None), not ratio_mode)
        self._set_widget_state(getattr(self, "_entry_runin_tlife", None), not ratio_mode)
        self._set_widget_state(getattr(self, "_entry_target_tlife", None), not ratio_mode)
        self._update_runin_tlife_labels()

    def _update_f_mech_label(self):
        try:
            rpm_min = float(self._vars["rpm_min"].get().strip() or 300.0)
            rpm_max = float(self._vars["rpm_max"].get().strip() or 300.0)
            if rpm_min > rpm_max:
                rpm_min, rpm_max = rpm_max, rpm_min
            m = int(float(self._vars["mech_harmonic"].get().strip() or 1))
            m = max(1, m)
            f_min = (rpm_min / 60.0) * m
            f_max = (rpm_max / 60.0) * m
            q_min = max(15.0, min(80.0, rpm_min / 10.0))
            q_max = max(15.0, min(80.0, rpm_max / 10.0))
            if abs(rpm_max - rpm_min) < 1e-12:
                self.f_mech_var.set(f"f_mech = rpm/60*m = {f_max:.6g} Hz，Notch Q≈{q_max:.3g}")
            else:
                self.f_mech_var.set(
                    f"f_mech = rpm/60*m = [{f_min:.6g}, {f_max:.6g}] Hz，Notch Q≈[{q_min:.3g}, {q_max:.3g}]"
                )
        except Exception:
            self.f_mech_var.set("f_mech=— Hz（输入有误）")


    def _update_ratio_labels(self):
        r1 = float(self.phase_runin_pct.get())
        r2 = float(self.phase_stable_pct.get())
        r3 = float(self.phase_severe_pct.get())
        s = max(1e-9, r1 + r2 + r3)
        r1, r2, r3 = (100.0 * r1 / s, 100.0 * r2 / s, 100.0 * r3 / s)
        self._lbl_runin.set(f"{r1:.2f}")
        self._lbl_stable.set(f"{r2:.2f}")
        self._lbl_severe.set(f"{r3:.2f}")
        self.phase_sum_label.set(f"磨合/稳定/加速：{r1:.2f}% / {r2:.2f}% / {r3:.2f}%（Σ=100%）")

    def _on_ratio_change(self, changed: str):
        """保持三段比例和为 100%。支持锁定一个阶段；未锁定时尽量保留另外两段的相对比例。"""
        if getattr(self, "phase_control_mode", None) is not None and self.phase_control_mode.get() != "ratio":
            self._update_ratio_labels()
            return
        if getattr(self, "_ratio_lock", False):
            return
        self._ratio_lock = True
        try:
            lock = self.locked_phase.get() if hasattr(self, "locked_phase") else "none"
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

            if lock not in ("none", "runin", "stable", "severe"):
                lock = "none"

            if lock == "none":
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
                else:
                    rem = 100.0 - r3
                    tot = r1 + r2
                    if tot <= 1e-9:
                        r1 = rem * 0.5
                        r2 = rem - r1
                    else:
                        r1 = rem * (r1 / tot)
                        r2 = rem - r1
                s = max(1e-9, r1 + r2 + r3)
                r1, r2, r3 = (100.0 * r1 / s, 100.0 * r2 / s, 100.0 * r3 / s)
            else:
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
                else:
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

                s = r1 + r2 + r3
                if abs(s - 100.0) > 1e-6:
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
        if getattr(self, "phase_control_mode", None) is not None and self.phase_control_mode.get() != "ratio":
            self._update_ratio_labels()
            return
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
            self._update_ratio_labels()
            return
        self._on_ratio_change(which)

    def _on_lock_change(self):
        """锁定一个阶段：禁用对应滑块与输入框"""
        lock = self.locked_phase.get() if hasattr(self, "locked_phase") else "none"

        if getattr(self, "phase_control_mode", None) is not None and self.phase_control_mode.get() != "ratio":
            for w in [
                getattr(self, "_scale_runin", None),
                getattr(self, "_entry_runin", None),
                getattr(self, "_scale_stable", None),
                getattr(self, "_entry_stable", None),
                getattr(self, "_scale_severe", None),
                getattr(self, "_entry_severe", None),
            ]:
                self._set_widget_state(w, False)
            return

        self._set_widget_state(getattr(self, "_scale_runin", None), lock != "runin")
        self._set_widget_state(getattr(self, "_entry_runin", None), lock != "runin")
        self._set_widget_state(getattr(self, "_scale_stable", None), lock != "stable")
        self._set_widget_state(getattr(self, "_entry_stable", None), lock != "stable")
        self._set_widget_state(getattr(self, "_scale_severe", None), lock != "severe")
        self._set_widget_state(getattr(self, "_entry_severe", None), lock != "severe")

        # 切换锁定后，做一次归一化
        self._on_ratio_change("runin")

    def _build_run_panel(self, parent):
        box = ttk.LabelFrame(parent, text="导出与输出", padding=10)
        box.pack(fill="x")

        ttk.Label(box, text="输出目录").grid(row=0, column=0, sticky="w")
        ttk.Entry(box, textvariable=self.out_dir, width=34).grid(row=1, column=0, sticky="we", pady=(4, 6))
        ttk.Button(box, text="选择…", command=self._choose_dir).grid(row=1, column=1, padx=(6, 0))

        ttk.Label(box, text="随机种子 seed").grid(row=2, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(box, textvariable=self.seed, width=12).grid(row=3, column=0, sticky="w", pady=(4, 6))

        # 开环-闭环对比输出范围
        ttk.Label(box, text="开环-闭环对比输出范围 (s)，(0,-1)=全部").grid(row=2, column=1, sticky="w", pady=(6, 0))
        rng_box = ttk.Frame(box)
        rng_box.grid(row=3, column=1, sticky="w", pady=(4, 6))
        ttk.Label(rng_box, text="t_start").pack(side="left")
        ttk.Entry(rng_box, textvariable=self.compare_t_start_s, width=8).pack(side="left", padx=(6, 10))
        ttk.Label(rng_box, text="t_end").pack(side="left")
        ttk.Entry(rng_box, textvariable=self.compare_t_end_s, width=8).pack(side="left", padx=(6, 0))

        # 图片语言选择
        lang_box = ttk.Frame(box)
        lang_box.grid(row=4, column=0, columnspan=2, sticky="we", pady=(6, 2))
        lang_box.columnconfigure(3, weight=1)

        ttk.Label(lang_box, text="图片语言").grid(row=0, column=0, sticky="w")
        ttk.Radiobutton(lang_box, text="中文", variable=self.plot_lang, value="zh").grid(row=0, column=1, sticky="w", padx=(10, 0))
        ttk.Radiobutton(lang_box, text="英文", variable=self.plot_lang, value="en").grid(row=0, column=2, sticky="w", padx=(10, 0))

        right_btns = ttk.Frame(lang_box)
        right_btns.grid(row=0, column=4, sticky="e")
        ttk.Button(right_btns, text="恢复默认值", command=self._restore_defaults).pack(side="left", padx=(0, 10))
        ttk.Button(right_btns, text="重置默认值", command=self._reset_defaults).pack(side="left", padx=(0, 10))
        ttk.Button(right_btns, text="保存默认值", command=self._save_defaults).pack(side="left")
# 两个按钮：导出数据 / 导出 图片
        self.btn_xlsx = ttk.Button(box, text="导出数据", command=self._start_export_data)
        self.btn_xlsx.grid(row=5, column=0, sticky="we", pady=(8, 4))
        self.btn_plots = ttk.Button(box, text="导出 图片", command=self._start_export_plots)
        self.btn_plots.grid(row=5, column=1, sticky="we", pady=(8, 4), padx=(6, 0))

        # 进度条（确定型） + 百分比
        self.pb = ttk.Progressbar(box, mode="determinate", maximum=100.0)
        self.pb.grid(row=6, column=0, sticky="we", pady=(6, 3))
        self.pb_pct = tk.StringVar(value="0%")
        ttk.Label(box, textvariable=self.pb_pct).grid(row=6, column=1, padx=(6, 0), sticky="w")

        # 底部区域：提示 + 日志（提示固定显示，日志可滚动）
        bottom = ttk.Frame(parent)
        bottom.pack(fill="both", expand=True, pady=(10, 0))
        bottom.columnconfigure(0, weight=1)
        bottom.rowconfigure(1, weight=1)

        tips = (
            "提示：\n"
            "1) 建议安装 SciPy：pip install scipy（陷波/低通更快更标准）\n"
            "2) duration 很长（例如10小时@50Hz）时，xlsx 会很大，可把 stride 设为 2~10。\n"
            "3) 若中文图仍乱码：请安装中文字体（微软雅黑/黑体/Noto Sans CJK/文泉驿等）。\n"
        )
        self.tips_label = ttk.Label(bottom, text=tips, justify="left", wraplength=560)
        self.tips_label.grid(row=0, column=0, sticky="we", pady=(0, 8))

        log_frame = ttk.Frame(bottom)
        log_frame.grid(row=1, column=0, sticky="nsew")
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        self.log = tk.Text(log_frame, height=12, width=54, wrap="word")
        ysb = ttk.Scrollbar(log_frame, orient="vertical", command=self.log.yview)
        self.log.configure(yscrollcommand=ysb.set)
        self.log.grid(row=0, column=0, sticky="nsew")
        ysb.grid(row=0, column=1, sticky="ns")

        self.log.insert("end", "日志输出：\n")
        self.log.configure(state="disabled")

        # 根据窗口宽度动态调整提示文本换行宽度
        def _on_bottom_resize(_evt):
            try:
                w = max(200, int(bottom.winfo_width()) - 20)
                self.tips_label.configure(wraplength=w)
            except Exception:
                pass
        bottom.bind("<Configure>", _on_bottom_resize)

    def _init_ini_paths(self):
        """初始化 ini 路径（兼容 PyInstaller onefile）。"""
        def _app_dir():
            if getattr(sys, "frozen", False):
                return Path(sys.executable).resolve().parent
            return Path(__file__).resolve().parent

        def _resource_dir():
            return Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))

        self._app_dir = _app_dir()
        self._resource_dir = _resource_dir()
        self._ini_path = self._app_dir / "defaultData.ini"
        self._ini_template_path = self._resource_dir / "defaultData.ini"

    def _fallback_ini_path(self) -> Path:
        base = Path.home() / "needle_hook_wear_sim"
        base.mkdir(parents=True, exist_ok=True)
        return base / "defaultData.ini"

    def _load_or_create_default_ini(self):
        """启动时读取 defaultData.ini；若不存在则生成。"""
        def _try_load_or_create():
            if self._ini_path.exists():
                self._read_default_ini()
                return
            # 优先从打包资源复制
            if hasattr(self, "_ini_template_path") and self._ini_template_path.exists():
                try:
                    self._ini_path.parent.mkdir(parents=True, exist_ok=True)
                    self._ini_path.write_bytes(self._ini_template_path.read_bytes())
                    self._read_default_ini()
                    return
                except Exception:
                    pass
            self._write_default_ini()

        try:
            _try_load_or_create()
        except Exception as e:
            try:
                # 尝试切换到用户目录
                self._ini_path = self._fallback_ini_path()
                _try_load_or_create()
                self._log(f"默认参数文件已切换到：{self._ini_path}\n")
            except Exception as e2:
                self._log(f"读取 defaultData.ini 失败：{e2}\n")

    def _read_default_ini(self):
        """读取 defaultData.ini（UTF-8），将默认值回填到 GUI。"""
        cp = configparser.ConfigParser(interpolation=None, strict=False)
        cp.read(self._ini_path, encoding="utf-8")
        sec = cp["DEFAULT"]

        def gs(k, default=""):
            return sec.get(k, fallback=str(default))

        self.out_dir.set(gs("out_dir", self.out_dir.get()))
        try:
            self.seed.set(int(float(gs("seed", self.seed.get()))))
        except Exception:
            pass
        self.plot_lang.set(gs("plot_lang", self.plot_lang.get()))
        self.data_export_format.set(gs("data_export_format", self.data_export_format.get()))
        self._set_export_tables_from_csv(gs("data_export_tables", self._export_tables_to_ini_value()))
        self._set_export_columns_from_csv(gs("data_export_columns", self._export_columns_to_ini_value()))
        self.invalid_ratio_pct.set(gs("invalid_ratio_pct", self.invalid_ratio_pct.get()))
        self.phase_control_mode.set(gs("phase_control_mode", self.phase_control_mode.get()))
        try:
            self.phase_runin_tlife_pct.set(float(gs("runin_ratio_by_tlife_pct", self.phase_runin_tlife_pct.get())))
        except Exception:
            pass
        self.target_tlife_s.set(gs("target_tlife_s", self.target_tlife_s.get()))
        self.compare_t_start_s.set(gs("compare_t_start_s", self.compare_t_start_s.get()))
        self.compare_t_end_s.set(gs("compare_t_end_s", self.compare_t_end_s.get()))

        # duration 以小时保存
        try:
            dur_h = float(gs("duration_h", 10.0))
            if "duration_h" in self._vars:
                self._vars["duration_h"].set(f"{dur_h}")
        except Exception:
            pass

        for k, var in self._vars.items():
            if k == "duration_h":
                continue
            if k in sec:
                var.set(gs(k, var.get()))

        # 兼容旧版：仅有 rpm 时，回填到 rpm_min/rpm_max
        if ("rpm_min" in self._vars) and ("rpm_max" in self._vars):
            if ("rpm_min" not in sec) and ("rpm_max" not in sec) and ("rpm" in sec):
                v = gs("rpm", "300")
                self._vars["rpm_min"].set(v)
                self._vars["rpm_max"].set(v)
            elif ("rpm_min" in sec) and ("rpm_max" not in sec):
                v = gs("rpm_min", "300")
                self._vars["rpm_max"].set(v)
            elif ("rpm_max" in sec) and ("rpm_min" not in sec):
                v = gs("rpm_max", "300")
                self._vars["rpm_min"].set(v)

        try:
            self.phase_runin_pct.set(float(gs("phase_runin_pct", self.phase_runin_pct.get())))
            self.phase_stable_pct.set(float(gs("phase_stable_pct", self.phase_stable_pct.get())))
            self.phase_severe_pct.set(float(gs("phase_severe_pct", self.phase_severe_pct.get())))
            self._on_ratio_change("runin")
        except Exception:
            pass
        self._update_runin_tlife_labels()

        if hasattr(self, "locked_phase") and ("locked_phase" in sec):
            self.locked_phase.set(gs("locked_phase", "none"))
        self._update_phase_mode_state()


        self._log("已读取 defaultData.ini 默认参数。\n")

    def _write_default_ini(self):
        """写入 defaultData.ini（UTF-8，含中文注释）。"""
        try:
            self._read_cfg()
        except Exception:
            pass

        try:
            dur_h = float(self._vars["duration_h"].get())
        except Exception:
            dur_h = self.cfg.duration_s / 3600.0

        items = [
            ("out_dir", "输出目录（导出 xlsx/图片/summary 的根目录）", self.out_dir.get()),
            ("seed", "随机种子（保证仿真可重复）", str(self.seed.get())),
            ("plot_lang", "导出图片语言：zh=中文，en=英文", self.plot_lang.get()),
            ("duration_h", "采样时间（小时，仅用于生成时间轴）", f"{dur_h}"),
            ("compare_t_start_s", "开环-闭环对比输出范围起始时间（秒；0=从头）", self.compare_t_start_s.get()),
            ("compare_t_end_s", "开环-闭环对比输出范围终止时间（秒；-1=到末尾）", self.compare_t_end_s.get()),
            ("locked_phase", "阶段比例锁定：none/runin/stable/severe", self.locked_phase.get()),
            ("phase_runin_pct", "磨合阶段比例（%）", f"{float(self.phase_runin_pct.get()):.6f}"),
            ("phase_stable_pct", "稳定磨损阶段比例（%）", f"{float(self.phase_stable_pct.get()):.6f}"),
            ("phase_severe_pct", "加速磨损阶段比例（%）", f"{float(self.phase_severe_pct.get()):.6f}"),
        ]
        items.extend([
            ("phase_control_mode", "阶段控制方式：ratio 或 runin_tlife", self.phase_control_mode.get()),
            ("runin_ratio_by_tlife_pct", "磨合比例 + tlife 控制的磨合段比例（%）", f"{float(self.phase_runin_tlife_pct.get()):.6f}"),
            ("target_tlife_s", "磨合比例 + tlife 控制的目标 tlife（s）", self.target_tlife_s.get()),
            ("invalid_ratio_pct", "invalid 比例（%，quality flag=0）", self.invalid_ratio_pct.get()),
            ("data_export_format", "数据导出格式：xlsx 或 db(SQLite)", self.data_export_format.get()),
            ("data_export_tables", "数据导出表（逗号分隔）", self._export_tables_to_ini_value()),
            ("data_export_columns", "数据导出列（逗号分隔）", self._export_columns_to_ini_value()),
        ])

        sim_comments = {
            "theta_deg": "包角 θ（单位：度）",
            "t_set_N": "平均张力设定值 T_set（单位：N）",
            "fs_Hz": "采样率 fs（Hz，仅用于生成时间轴）",
            "rpm_min": "转速范围下限 rpm_min",
            "rpm_max": "转速范围上限 rpm_max",
            "mech_harmonic": "机械扰动谐波阶次 m（主频= rpm/60*m）",
            "tau_rpm_s": "转速慢变时间常数 τ_rpm（s，<=0关闭）",
            "noise_rms_open_min": "开环：高频噪声范围下限（RMS，N）",
            "noise_rms_open_max": "开环：高频噪声范围上限（RMS，N）",
            "noise_rms_closed_min": "闭环：高频噪声范围下限（RMS，N）",
            "noise_rms_closed_max": "闭环：高频噪声范围上限（RMS，N）",
            "mech_amp_open_min": "开环：机械周期扰动幅值范围下限（N）",
            "mech_amp_open_max": "开环：机械周期扰动幅值范围上限（N）",
            "mech_amp_closed_min": "闭环：机械周期扰动幅值范围下限（N）",
            "mech_amp_closed_max": "闭环：机械周期扰动幅值范围上限（N）",
            "drift_amp_open_min": "开环：低频漂移幅值范围下限（N）",
            "drift_amp_open_max": "开环：低频漂移幅值范围上限（N）",
            "drift_amp_closed_min": "闭环：低频漂移幅值范围下限（N）",
            "drift_amp_closed_max": "闭环：低频漂移幅值范围上限（N）",
            "drift_freq_hz_min": "低频漂移频率范围下限（Hz）",
            "drift_freq_hz_max": "低频漂移频率范围上限（Hz）",
            "sensor_rms_min": "张力测量噪声范围下限（RMS，N）",
            "sensor_rms_max": "张力测量噪声范围上限（RMS，N）",
            "tau_mech_s": "机械周期扰动幅值慢变时间常数 τ_mech（s，<=0关闭慢变）",
            "tau_noise_s": "高频噪声 RMS 慢变时间常数 τ_noise（s，<=0关闭慢变）",
            "tau_sensor_s": "传感器测量噪声 RMS 慢变时间常数 τ_sensor（s，<=0关闭慢变）",
            "tau_drift_amp_s": "低频漂移幅值慢变时间常数 τ_driftA（s，<=0关闭慢变）",
            "tau_drift_freq_s": "低频漂移频率慢变时间常数 τ_driftf（s，<=0关闭慢变）",
            "tmin_gate_N": "张力门控下限（N，避免比值/对数放大）",
            "ratio_clip_min": "R=T2/T1 比值裁剪下限",
            "ratio_clip_max": "R=T2/T1 比值裁剪上限",
            "hampel_win_s": "Hampel 去毛刺窗口长度（s）",
            "hampel_nsig": "Hampel 判别阈值（倍标准差）",
            "lowpass_fc_hz": "低通滤波截止频率（Hz）",
            "stable_win_s": "稳定段窗口长度 W_ss（s）",
            "stable_hold_s": "最短连续稳定段 Whold（s）",
            "stable_sigma_max": "稳定段标准差阈值 σ_max",
            "stable_slope_max": "稳定段斜率阈值 g_max（绝对值）",
            "stable_valid_min": "稳定段有效样本比例下限 Neff/Nss",
            "fail_delta": "失效阈值相对增量 δ（μth=μss*(1+δ)）",
            "fail_hold_s": "超限保持时间 Wpersist（s）",
            "export_stride": "导出步长 stride（=1 全部导出；>1 抽样导出）",
            "plot_max_points": "绘图最大点数（超过则自动降采样）",
            "mu_runin_min": "磨合段 μ 范围下限",
            "mu_runin_max": "磨合段 μ 范围上限",
            "mu_stable_min": "稳定段 μ 范围下限",
            "mu_stable_max": "稳定段 μ 范围上限",
            "mu_severe_min": "加速段 μ 范围下限",
            "mu_severe_max": "加速段 μ 范围上限",
            "trans_runin2stable_k": "磨合→稳定过渡速度系数 k_rs（越大越快）",
            "trans_stable2severe_k": "稳定→加速过渡速度系数 k_sa（越大越快）",
        }

        for k, cmt in sim_comments.items():
            try:
                v = getattr(self.cfg, k)
            except Exception:
                continue
            items.append((k, cmt, str(v)))

        lines = []
        lines.append("# -*- coding: utf-8 -*-")
        lines.append("; defaultData.ini 自动生成：用于保存/读取 GUI 默认参数（UTF-8）")
        lines.append("; 修改后下次启动会自动加载。")
        lines.append("[DEFAULT]")
        seen = set()
        for k, cmt, v in items:
            kk = str(k).strip().lower()
            if kk in seen:
                continue
            seen.add(kk)
            lines.append(f"; {cmt}")
            lines.append(f"{k}={v}")
            lines.append("")
        self._ini_path.parent.mkdir(parents=True, exist_ok=True)
        self._ini_path.write_text("\n".join(lines), encoding="utf-8")
        self._log("已写入 defaultData.ini 默认参数。\n")

    def _write_default_ini_factory(self):
        """将 defaultData.ini 重置为“程序内置默认值”（UTF-8，含中文注释）。"""
        from needle_hook_wear_simulator_gui import SimConfig
        cfg0 = SimConfig()
        try:
            cfg0.validate()
        except Exception:
            pass

        # GUI 层内置默认值（不要用当前 GUI 输入，避免“改参数就写 ini”）
        out_dir0 = os.path.abspath("sim_out")
        seed0 = 7
        plot_lang0 = "zh"
        data_export_format0 = "xlsx"
        data_export_tables0 = ",".join(DEFAULT_EXPORT_TABLES)
        data_export_columns0 = ",".join(DEFAULT_EXPORT_COLUMNS)
        compare_t_start0 = 0.0
        compare_t_end0 = -1.0
        locked_phase0 = "none"
        phase_control_mode0 = str(getattr(cfg0, "phase_control_mode", "ratio"))
        runin_ratio_by_tlife_pct0 = float(getattr(cfg0, "runin_ratio_by_tlife", cfg0.phase_runin_ratio)) * 100.0
        target_tlife_s0 = float(getattr(cfg0, "target_tlife_s", cfg0.duration_s * (cfg0.phase_runin_ratio + cfg0.phase_stable_ratio)))
        invalid_ratio_pct0 = float(getattr(cfg0, "invalid_ratio", 0.0005)) * 100.0

        # 阶段比例（%）
        phase_runin_pct0 = float(cfg0.phase_runin_ratio) * 100.0
        phase_stable_pct0 = float(cfg0.phase_stable_ratio) * 100.0
        phase_severe_pct0 = float(cfg0.phase_severe_ratio) * 100.0

        dur_h0 = float(cfg0.duration_s) / 3600.0
        factory_export_items = [
            ("phase_control_mode", "阶段控制方式：ratio 或 runin_tlife", str(phase_control_mode0)),
            ("runin_ratio_by_tlife_pct", "磨合比例 + tlife 控制的磨合段比例（%）", f"{runin_ratio_by_tlife_pct0:.6f}"),
            ("target_tlife_s", "磨合比例 + tlife 控制的目标 tlife（s）", f"{target_tlife_s0:.6f}"),
            ("invalid_ratio_pct", "invalid 比例（%，quality flag=0）", f"{invalid_ratio_pct0:.6f}"),
            ("data_export_format", "数据导出格式：xlsx 或 db(SQLite)", str(data_export_format0)),
            ("data_export_tables", "数据导出表（逗号分隔）", str(data_export_tables0)),
            ("data_export_columns", "数据导出列（逗号分隔）", str(data_export_columns0)),
        ]

        items = factory_export_items + [
            ("out_dir", "输出目录（导出 xlsx/图片/summary 的根目录）", str(out_dir0)),
            ("seed", "随机种子（保证仿真可重复）", str(seed0)),
            ("plot_lang", "导出图片语言：zh=中文，en=英文", str(plot_lang0)),
            ("duration_h", "采样时间（小时，仅用于生成时间轴）", f"{dur_h0}"),
            ("compare_t_start_s", "开环-闭环对比输出范围起始时间（秒；0=从头）", str(compare_t_start0)),
            ("compare_t_end_s", "开环-闭环对比输出范围终止时间（秒；-1=到末尾）", str(compare_t_end0)),
            ("locked_phase", "阶段比例锁定：none/runin/stable/severe", str(locked_phase0)),
            ("phase_runin_pct", "磨合阶段比例（%）", f"{phase_runin_pct0:.6f}"),
            ("phase_stable_pct", "稳定磨损阶段比例（%）", f"{phase_stable_pct0:.6f}"),
            ("phase_severe_pct", "加速磨损阶段比例（%）", f"{phase_severe_pct0:.6f}"),
        ]

        sim_comments = {
            "theta_deg": "包角 θ（单位：度）",
            "t_set_N": "平均张力设定值 T_set（单位：N）",
            "fs_Hz": "采样率 fs（Hz，仅用于生成时间轴）",
            "duration_s": "采样时间（秒，内部字段；GUI 使用 duration_h 显示）",
            "rpm_min": "转速范围下限 rpm_min",
            "rpm_max": "转速范围上限 rpm_max",
            "mech_harmonic": "机械扰动谐波阶次 m（主频= rpm/60*m）",
            "tau_rpm_s": "转速慢变时间常数 τ_rpm（s，<=0关闭）",
            "noise_rms_open_min": "开环：高频噪声范围下限（RMS，N）",
            "noise_rms_open_max": "开环：高频噪声范围上限（RMS，N）",
            "noise_rms_closed_min": "闭环：高频噪声范围下限（RMS，N）",
            "noise_rms_closed_max": "闭环：高频噪声范围上限（RMS，N）",
            "mech_amp_open_min": "开环：机械周期扰动幅值范围下限（N）",
            "mech_amp_open_max": "开环：机械周期扰动幅值范围上限（N）",
            "mech_amp_closed_min": "闭环：机械周期扰动幅值范围下限（N）",
            "mech_amp_closed_max": "闭环：机械周期扰动幅值范围上限（N）",
            "drift_amp_open_min": "开环：低频漂移幅值范围下限（N）",
            "drift_amp_open_max": "开环：低频漂移幅值范围上限（N）",
            "drift_amp_closed_min": "闭环：低频漂移幅值范围下限（N）",
            "drift_amp_closed_max": "闭环：低频漂移幅值范围上限（N）",
            "drift_freq_hz_min": "低频漂移频率范围下限（Hz）",
            "drift_freq_hz_max": "低频漂移频率范围上限（Hz）",
            "sensor_rms_min": "张力测量噪声范围下限（RMS，N）",
            "sensor_rms_max": "张力测量噪声范围上限（RMS，N）",
            "tmin_gate_N": "张力门控下限（N，避免比值/对数放大）",
            "ratio_clip_min": "R=T2/T1 比值裁剪下限",
            "ratio_clip_max": "R=T2/T1 比值裁剪上限",
            "hampel_win_s": "Hampel 去毛刺窗口长度（s）",
            "hampel_nsig": "Hampel 判别阈值（倍标准差）",
            "lowpass_fc_hz": "低通滤波截止频率（Hz）",
            "stable_win_s": "稳定段窗口长度 W_ss（s）",
            "stable_hold_s": "最短连续稳定段 Whold（s）",
            "stable_sigma_max": "稳定段标准差阈值 σ_max",
            "stable_slope_max": "稳定段斜率阈值 g_max（绝对值）",
            "stable_valid_min": "稳定段有效样本比例下限 Neff/Nss",
            "fail_delta": "失效阈值相对增量 δ（μth=μss*(1+δ)）",
            "fail_hold_s": "超限保持时间 Wpersist（s）",
            "export_stride": "导出步长 stride（=1 全部导出；>1 抽样导出）",
            "plot_max_points": "绘图最大点数（超过则自动降采样）",
            "mu_runin_min": "磨合段 μ 范围下限",
            "mu_runin_max": "磨合段 μ 范围上限",
            "mu_stable_min": "稳定段 μ 范围下限",
            "mu_stable_max": "稳定段 μ 范围上限",
            "mu_severe_min": "加速段 μ 范围下限",
            "mu_severe_max": "加速段 μ 范围上限",
            "trans_runin2stable_k": "磨合→稳定过渡系数 k_rs（越大越快）",
            "trans_stable2severe_k": "稳定→加速过渡系数 k_sa（越大越快）",
}

        for k, cmt in sim_comments.items():
            try:
                v = getattr(cfg0, k)
            except Exception:
                continue
            items.append((k, cmt, str(v)))

        lines = []
        lines.append("# -*- coding: utf-8 -*-")
        lines.append("; defaultData.ini 自动生成：用于保存/读取 GUI 默认参数（UTF-8）")
        lines.append("; 说明：本文件仅在“保存默认值/重置默认值”时被改写。")
        lines.append("[DEFAULT]")
        seen = set()
        for k, cmt, v in items:
            kk = str(k).strip().lower()
            if kk in seen:
                continue
            seen.add(kk)
            lines.append(f"; {cmt}")
            lines.append(f"{k}={v}")
            lines.append("")
        self._ini_path.parent.mkdir(parents=True, exist_ok=True)
        self._ini_path.write_text("\n".join(lines), encoding="utf-8")

    def _on_close(self):
        """（兼容保留）关闭窗口：默认不自动改写 defaultData.ini。"""
        self.destroy()

    def _save_defaults(self):
        """手动保存当前参数为 defaultData.ini（不会在关闭/导出时自动写入）。"""
        try:
            self._write_default_ini()
            self._log("已手动保存 defaultData.ini 默认参数。\n")
        except Exception as e:
            self._log(f"手动保存 defaultData.ini 失败：{e}\n")

    def _restore_defaults(self):
        """恢复默认值：从 defaultData.ini 读取并回填到 GUI（不写入 ini）。"""
        try:
            if hasattr(self, "_ini_path") and self._ini_path.exists():
                self._read_default_ini()
                self._log("已从 defaultData.ini 恢复默认参数。\n")
            else:
                self._log("未找到 defaultData.ini，无法恢复（可先点击“保存默认值”生成）。\n")
        except Exception as e:
            self._log(f"从 defaultData.ini 恢复失败：{e}\n")



    def _reset_defaults(self):
        """重置默认值：重置 defaultData.ini 为“程序内置默认值”，并将 GUI 输入同步重置。"""
        try:
            self._write_default_ini_factory()
            self._read_default_ini()
            self._log("已重置 defaultData.ini，并已重置 GUI 输入为默认值。\n")
        except Exception as e:
            self._log(f"重置默认值失败：{e}\n")


    def _choose_dir(self):
        d = filedialog.askdirectory(title="选择输出目录")
        if d:
            self.out_dir.set(d)

    
    def _parse_range_text(self, s: str, default: float = 0.0):
        """解析范围输入：支持 'a,b' 或 'a~b' 或单个数值；返回 (min,max)。"""
        if s is None:
            v = float(default)
            return v, v
        ss = str(s).strip().replace("～", "~")
        if ss == "":
            v = float(default)
            return v, v
        if "," in ss:
            parts = [p.strip() for p in ss.split(",") if p.strip() != ""]
        elif "~" in ss:
            parts = [p.strip() for p in ss.split("~") if p.strip() != ""]
        else:
            parts = [ss]
        try:
            if len(parts) == 1:
                v = float(parts[0])
                return v, v
            a = float(parts[0]); b = float(parts[1])
            if a > b:
                a, b = b, a
            return a, b
        except Exception:
            v = float(default)
            return v, v

    def _format_range(self, a: float, b: float) -> str:
        """将范围格式化为 'min,max' 字符串（GUI 显示用）。"""
        try:
            a = float(a); b = float(b)
        except Exception:
            return ""
        if abs(a - b) < 1e-12:
            return f"{a}"
        return f"{a},{b}"

    def _read_cfg(self):
        # 读取 GUI 输入并写入 cfg（扰动范围为两个输入框：min/max）
        for k, v in self._vars.items():
            s = v.get().strip()
            if s == "":
                continue
            try:
                typ = self._types.get(k, float)
                val = typ(float(s))
                if k == "duration_h":
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
        try:
            self.cfg.phase_control_mode = str(self.phase_control_mode.get()).strip() or "ratio"
        except Exception:
            self.cfg.phase_control_mode = "ratio"
        try:
            self.cfg.runin_ratio_by_tlife = float(self.phase_runin_tlife_pct.get()) / 100.0
        except Exception:
            self.cfg.runin_ratio_by_tlife = float(getattr(self.cfg, "phase_runin_ratio", 0.12))
        try:
            self.cfg.target_tlife_s = float(self.target_tlife_s.get().strip())
        except Exception:
            self.cfg.target_tlife_s = float(getattr(self.cfg, "duration_s", 0.0) * max(0.0, getattr(self.cfg, "phase_runin_ratio", 0.12)))
        try:
            self.cfg.invalid_ratio = float(self.invalid_ratio_pct.get().strip()) / 100.0
        except Exception:
            self.cfg.invalid_ratio = 0.0005

        # 开环-闭环对比输出范围
        try:
            self.cfg.compare_t_start_s = float(self.compare_t_start_s.get().strip())
        except Exception:
            self.cfg.compare_t_start_s = 0.0
        try:
            self.cfg.compare_t_end_s = float(self.compare_t_end_s.get().strip())
        except Exception:
            self.cfg.compare_t_end_s = -1.0

        self.cfg.validate()



    def _make_cache_key(self) -> str:
        """对会影响仿真“数据本体/判据结果”的参数做指纹，避免漏项导致误用缓存。
        注意：导出步长/绘图点数/对比输出范围等仅影响导出呈现，不必触发重算。
        """
        from dataclasses import asdict
        import json, hashlib
        cfgd = asdict(self.cfg)

        # 仅影响导出呈现的字段（不触发重算）
        for k in ["export_stride", "plot_max_points", "compare_t_start_s", "compare_t_end_s"]:
            if k in cfgd:
                cfgd.pop(k)

        payload = {"cfg": cfgd, "seed": int(self.seed.get())}
        s = json.dumps(payload, sort_keys=True, ensure_ascii=False, separators=(",", ":"))
        return hashlib.md5(s.encode("utf-8")).hexdigest()


    def _ensure_simulated(self, progress_cb, prepare_plot: bool = False):
        progress_cb(1.0, "开始生成仿真数据（用于导出）...")
        if prepare_plot:
            setup_plot_font(lang=self.plot_lang.get(), progress_cb=progress_cb, pct=2.0)
        key = self._make_cache_key()
        if self._cache_res is not None and self._cache_key == key:
            progress_cb(10.0, "参数未变，使用缓存的仿真结果（跳过重算）")
            return self._cache_res
        res = simulate(self.cfg, seed=int(self.seed.get()), progress_cb=progress_cb)
        self._cache_key = key
        self._cache_res = res
        return res

    def _start_export_data(self):
        self._start_task(task="data")

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

        selected_columns = None
        selected_tables = None
        task_label = task
        if task == "data":
            selected_tables = self._get_selected_export_tables()
            if not selected_tables:
                messagebox.showerror("输入错误", "请至少勾选一张导出表。")
                return
            selected_columns = self._get_selected_export_columns()
            if not selected_columns:
                messagebox.showerror("输入错误", "请至少勾选一列导出数据。")
                return
            task_label = f"data ({self.data_export_format.get()})"

        self.btn_xlsx.configure(state="disabled")
        self.btn_plots.configure(state="disabled")
        self.pb["value"] = 0.0
        self.pb_pct.set("0%")
        self.status.set("运行中…")
        self._log(f"开始任务：导出 {task_label}\n")

        def progress_cb(pct, msg):
            self.msg_q.put(("progress", float(pct), str(msg)))

        def work():
            try:
                res = self._ensure_simulated(progress_cb, prepare_plot=(task == "plots"))
                outputs = {}
                if task == "data":
                    fmt = str(self.data_export_format.get()).strip().lower() or "xlsx"
                    if fmt == "sqlite":
                        fmt = "db"
                    outputs[fmt] = export_data(
                        res,
                        out_dir=out_dir,
                        export_format=fmt,
                        selected_columns=selected_columns,
                        selected_tables=selected_tables,
                        progress_cb=progress_cb,
                    )
                elif task == "plots":
                    outputs.update(export_plots(
                        res,
                        out_dir=out_dir,
                        lang=self.plot_lang.get(),
                        progress_cb=progress_cb,
                        font_prepared=True,
                    ))
                export_summary(res, out_dir=out_dir, extra={"outputs": outputs})
                self.msg_q.put(("done", task_label))
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
