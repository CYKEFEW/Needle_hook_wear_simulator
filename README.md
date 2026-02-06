# 针钩磨损仿真（GUI）

基于 Tkinter 的针钩磨损过程仿真工具，支持参数配置、导出 xlsx 与图片，并生成 `summary.json` 结果汇总。

**功能**
1. 三阶段磨损模型（磨合/稳定/加速）与扰动建模（机械周期/噪声/漂移）。
2. 滤波与判据（Hampel、陷波、低通、稳定段与寿命判定）。
3. 导出 `xlsx`、图片与 `summary.json`。
4. 支持中文/英文绘图。

**运行环境**
1. 建议使用 Python 3.9 及以上版本。
2. Windows 已验证运行（需系统存在中文字体以避免图片乱码）。

**安装依赖**
```bash
pip install -r requirements.txt
```

可选安装（更标准的滤波实现）：
```bash
pip install scipy
```

**运行**
```bash
python needle_hook_wear_sim_gui_app.py
```

**导出说明**
1. 输出目录默认 `sim_out`。
2. 导出内容包括 `needle_hook_wear_sim.xlsx`、`plots/` 图片文件与 `summary.json`。

**默认参数**
1. 默认参数文件为 `defaultData.ini`。
2. 程序启动时会读取 `defaultData.ini`，不存在则自动生成。
3. GUI 内可“保存默认值/重置默认值”。

**打包（PyInstaller）**
1. 安装打包工具：
```bash
pip install pyinstaller
```
2. 使用本仓库提供的 `main.spec`：
```bash
pyinstaller main.spec
```
3. 产物在 `dist/needle_hook_wear_sim/` 目录。

**常见问题**
1. 导出图片中文乱码：安装微软雅黑/黑体/Noto Sans CJK 等中文字体。
2. 导出 xlsx 报错：确认已安装 `openpyxl` 与 `xlsxwriter`。
