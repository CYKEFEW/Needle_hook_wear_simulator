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
        tab_filter = ttk.Frame(nb, padding=10)
        tab_judge = ttk.Frame(nb, padding=10)
        tab_export = ttk.Frame(nb, padding=10)
        nb.add(tab_core, text="核心输入")
        nb.add(tab_dist, text="扰动参数")
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
        self.cfg.validate()

    def _make_cache_key(self) -> str:
        items = [
            ("theta", self.cfg.theta_deg),
            ("tset", self.cfg.t_set_N),
            ("fs", self.cfg.fs_Hz),
            ("dur_h", self.cfg.duration_s/3600.0),
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
