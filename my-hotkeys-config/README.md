# WPS / Excel / 浏览器 自定义快捷键

基于 AutoHotkey v2 的个人快捷键配置。脚本文件：`wps-excel-shortcuts.ahk`

## 环境

- Windows 11
- AutoHotkey v2.0.31（`winget install AutoHotkey.AutoHotkey`）
- WPS Office 12.1（表格进程为 `wps.exe`，窗口类 `OpusApp`）
- Microsoft Office 365（`EXCEL.EXE`）

## 快捷键总表

### 全机通用（任意程序）

| 快捷键 | 功能 |
|---|---|
| `Alt+1` | 粘贴 |
| `Alt+2` | 复制 |
| `Alt+3` | 选择性粘贴 |
| `` Alt+` `` / `Alt+~` | 全选 |

### 按程序区分（Alt+Q / Alt+W）

| 程序 | `Alt+Q` | `Alt+W` |
|---|---|---|
| 浏览器 | 下一个标签 | 上一个标签 |
| Q-Dir / Notion | 下一个标签 | 上一个标签 |
| WPS / Excel | 右移一格 | 左移一格 |
| CCHH | （未映射） | （未映射） |

### 原生即可用（未覆盖，保持原样）

- `Ctrl+Shift+L` 筛选
- `Ctrl+Shift+方向键 / Home / End / PgUp / PgDn` 扩展选区 / 切换工作表
- `Ctrl+Shift+~ ! @ # $ % ^ &` 数字格式（常规/千分位/时间/日期/货币/百分比/科学计数/外边框）
- `Ctrl+Shift+> / <` 增大 / 减小字号
- `Ctrl+Shift+=` / `Ctrl+-` 插入 / 删除单元格行

## 已知限制

- **CCHH（Claude Code Haha）**：源码 `!e.altKey` 写死，按住 Alt 时 Ctrl+Tab 不触发，故未做 Alt+Q/W 映射；使用其原生 `Ctrl+Tab` / `Ctrl+Shift+Tab`。
- **WPS**：按住 Alt 时切标签不稳定（与 CCHH 同理），故 Alt+Q/W 在 WPS 内映射为移动单元格。

## 安装 / 使用

1. 安装 AutoHotkey v2。
2. 双击 `wps-excel-shortcuts.ahk` 运行。
3. 开机自启：把脚本快捷方式放入
   `%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup`。

## 自检

`Ctrl+Alt+F12`：弹出当前前台程序名、窗口类名、Excel/表格识别状态。
