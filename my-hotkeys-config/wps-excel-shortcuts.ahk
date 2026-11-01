#Requires AutoHotkey v2.0
#SingleInstance Force
; ============================================================
; WPS 表格 + Microsoft Excel 自定义快捷键
; 作用域：仅当 Excel(EXCEL.EXE) 或 WPS 表格(et.exe) 窗口激活时生效
; 版本：2026-11-01 / 77
;
; 分级：
;   ✅ 稳定实现（纯快捷键映射，Excel/WPS 通用，已可直接用）
;   🔧 可用但需你测一下（依赖对话框默认焦点）
;   ⚠️ 待绑定（命令类，中文菜单加速键不可靠，需改用宏/UIA）
; ============================================================
#ErrorStdOut
InstallKeybdHook
SetKeyDelay 50, 25

; WPS 12 起表格可能跑在 wps.exe（统一进程），故一并纳入；后续按类名精修
IsSheet() => WinActive("ahk_exe EXCEL.EXE") || WinActive("ahk_exe et.exe") || WinActive("ahk_exe wps.exe")
IsExcel() => WinActive("ahk_exe EXCEL.EXE")

ShowTip(msg, ms := 1600) {
    ToolTip(msg, 20, 20, 3)
    SetTimer(() => ToolTip("", , , 3), -ms)
}
Pending(what) => ShowTip("⚠️ 待绑定：" what)

#HotIf IsSheet()

; ================= 自定义的字母 =================
; 1. Ctrl+Shift+S 发送文件  ⚠️（中文菜单不可靠）
^+s::Pending("发送文件")

; 2. Ctrl+Shift+D 编辑链接  ⚠️
^+d::Pending("编辑链接")

; 3. Ctrl+Shift+P 数据透视表  ⚠️（需菜单/宏绑定）
^+p::Pending("数据透视表")

; 4. Ctrl+Shift+W 任务窗格  ⚠️
^+w::Pending("任务窗格")

; 5. Ctrl+Shift+N 文本型数字→数字  🔧 Excel=分列(A,E 后回车)  WPS 需测
^+n::{
    if IsExcel() {
        Send "!ae"
        Sleep 120
        Send "{Enter}"
    } else Pending("文本转数字(WPS)")
}

; 6. Ctrl+Shift+T 数字→文本型数字  🔧 先把单元格设为文本格式(@)
^+t::Send "^1"   ; 打开格式对话框，可手动选文本；或下方改为直接设格式

; 7. Ctrl+3 打开字体窗口  🔧 Ctrl+1 打开「格式单元格」含字体页
^3::Send "^1"

; 8. Ctrl+Shift+H 粘贴值到可见单元格  ⚠️
^+h::Pending("粘贴值到可见单元格")

; 9~11. Alt+1/2/3 已移至下方「全机通用」区

; 12. Ctrl+Shift+Delete 删除单元格/行  ✅
^+Delete::Send "^-"

; 13. Ctrl+2 填充颜色  ⚠️（调色板无法用稳定快捷键唤起）
^2::Pending("填充颜色")

; ================= 系统设置的字母 =================
; Ctrl+Shift+L 筛选      → Excel/WPS 原生已有，保持不变
^+a::Send "^a"           ; Ctrl+Shift+A 全选 ✅
; Ctrl+Shift+V 仅粘贴为数值 🔧 Excel 选择性粘贴->V
^+v::{
    Send "^!v"
    Sleep 180
    Send "v"
    Sleep 80
    Send "{Enter}"
}

; ================= 标点 / 辅助 / 数字符号 =================
; 以下均为 Excel/WPS 原生已内置，直接可用，无需覆盖：
;   Ctrl+Shift+方向键/Home/End/PgUp/PgDn  → 扩展选区 / 切换工作表
;   Ctrl+Shift+~ ! @ # $ % ^ &           → 常规/千分位/时间/日期/货币/百分比/科学计数/外边框
;   Ctrl+Shift+> / <                      → 增大/减小字号
;   Ctrl+Shift+= / Ctrl+-                 → 插入/删除单元格行
;   Ctrl+Shift+{                          → 查找引用

; Ctrl+Shift+/ 显示当前时间 ✅（插入 HH:mm:ss 文本）
^+/::SendText FormatTime(, "HH:mm:ss")

#HotIf

; ================= 全机通用（不限窗口，任何软件都生效）=================
; Alt+1 粘贴
!1::Send "^v"
; Alt+2 复制
!2::Send "^c"
; Alt+3 选择性粘贴（在支持的应用里弹选择性粘贴，不支持则等同普通粘贴）
!3::Send "^!v"

; Alt+` 或 Alt+~ 全选（sc029 = Tab 上方那颗键；Alt+~ 是它的 Shift 形态）
!sc029::Send "^a"
!+sc029::Send "^a"

; ================= Alt+Q / Alt+W 键盘行为 =================
; 浏览器等：切标签（Ctrl+PgDn / Ctrl+PgUp）
; WPS/Excel：移动单元格光标（右 / 左）
!q::SwitchTab(1)
!w::SwitchTab(-1)

SwitchTab(dir) {
    proc := StrLower(WinGetProcessName("A"))
    isSheet := (proc = "wps.exe" || proc = "et.exe" || proc = "excel.exe")
    isTabApp := (proc = "notion.exe" || proc = "q-dir.exe")
    if (isSheet) {
        ; 表格类：发方向键，右移/左移单元格
        Send dir > 0 ? "{Right}" : "{Left}"
    } else if (isTabApp) {
        ; Q-Dir / Notion：优先按住 Alt 直接发 Ctrl+Tab
        Send dir > 0 ? "^{Tab}" : "^+{Tab}"
    } else {
        ; 其他（含浏览器）：Ctrl+PgDn / Ctrl+PgUp 切标签
        Send dir > 0 ? "^{PgDn}" : "^{PgUp}"
    }
}

; CCHH 不做自定义：用其原生 Ctrl+Tab（源码要求不按 Alt）

; ================= 自检（全局，任何时候都能按）=================
^!F12::{
    proc := WinGetProcessName("A")
    title := WinGetTitle("A")
    cls := WinGetClass("A")
    s := "前台程序：" proc "`n窗口类名：" cls "`n标题：" title
      . "`n`nExcel 识别：" (WinActive("ahk_exe EXCEL.EXE") ? "是" : "否")
      . "`n快捷键生效：" (IsSheet() ? "是 ✅" : "否 ❌")
    MsgBox s, "自检"
}
TrayTip "WPS/Excel 快捷键已就绪", "仅在表格窗口生效；Ctrl+Alt+F12 自检", 1
