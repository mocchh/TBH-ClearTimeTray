# TBH 通关时间监控（托盘）

系统托盘小程序：通过 Frida 读取 *Taskbar Hero* 通关通知，按关卡记录**最近 10 次**通关秒数，写入 Excel 并自动计算平均通关时间。

## 功能

- 系统托盘常驻，自动附加 `TaskBarHero` 进程
- Hook `TMP_Text.SetText`，解析通关富文本通知
- 按 **难度 + 关卡** 分别统计最近 10 次秒数（普通 / 噩梦 / 地狱 / 折磨）
- 生成 `data/clear_times.xlsx`（汇总 + 明细 + 平均）
- 进程退出后自动重连

## 通关文案格式

界面显示（**不含难度名**）：

```text
通关了关卡 2-9 (199秒) [11:32]
```

难度从游戏内存 `UI_Portal.m_currentStageDifficulty` 读取：

| 值 | 难度 |
|----|------|
| 0 | 普通 |
| 1 | 噩梦 |
| 2 | 地狱 |
| 3 | 折磨 |

Excel / 统计主键形如：`折磨|2-9`（与 `普通|2-9` 分开平均）。

## 环境

- Windows 10/11 x64
- Python 3.10+（源码运行）
- 建议**管理员权限**运行，以便 Frida 附加游戏

## 安装依赖

```powershell
cd TBH-ClearTimeTray
pip install -r requirements.txt
```

## 源码运行

```powershell
python -u tray_app.py
```

或双击 `run.bat`。

## 打包 EXE

```powershell
.\build_exe.bat
```

产物：`dist\TBH通关时间监控.exe`

也可：

```powershell
pip install pyinstaller
pyinstaller --noconfirm --clean --onefile --windowed --name "TBH通关时间监控" --add-data "clear_time_probe.js;." --hidden-import excel_store --collect-all frida --collect-all pystray --collect-all openpyxl tray_app.py
```

## 托盘菜单

| 菜单 | 说明 |
|------|------|
| 系统通知（默认关） | 开关通关气泡通知，**默认关闭** |
| 提示音（默认关） | 开关提示音，**默认关闭** |
| 打开 Excel | 打开 `data/clear_times.xlsx` |
| 打开数据目录 | 打开 `data/` |
| 重新连接游戏 | 重新附加进程 |
| 退出 | 结束程序 |

数据目录位于 **exe 同级**（或源码目录）下的 `data/`。

### 持久化

- 主数据：`data/clear_times.json`（关闭软件不会清空）
- 可读表：`data/clear_times.xlsx`（每次记录同步；启动时从 JSON 恢复）
- 配置：`data/config.json`（通知开关等）

### DPI

启动时启用 Per-Monitor DPI Aware，托盘图标按系统缩放生成，菜单跟随 Windows 缩放。

## Excel 结构

**汇总**

| 难度 | 关卡 | 第1次…第10次(秒) | 样本数 | 平均通关(秒) | 最近记录时间 | 最近通知时钟 |
|------|------|------------------|--------|--------------|--------------|--------------|

**明细**：难度 + 关卡 + 秒数；各组合最近最多 10 条。

## 目录

```text
TBH-ClearTimeTray/
├── tray_app.py
├── excel_store.py
├── clear_time_probe.js
├── build_exe.bat
├── run.bat
├── requirements.txt
└── data/                 # 运行后生成
    ├── clear_times.xlsx
    └── tray.log
```

## 免责声明

本项目仅供学习与研究。使用 Frida 注入、内存读取可能违反游戏用户协议。请自行承担使用风险；作者不对账号、数据或系统损失负责。

## License

MIT
