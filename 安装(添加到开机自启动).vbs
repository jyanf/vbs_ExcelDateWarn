Option Explicit
' On Error Resume Next
dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
dim ws : Set ws = CreateObject("WScript.Shell")
dim path : path = fso.GetFolder(".").Path & "\"
dim file : Set file = fso.GetFile(path & "日期检查.vbs")
dim tg : tg = ws.SpecialFolders("Startup") & "\"

' file.Copy tg
dim sc
Set sc = ws.CreateShortcut(tg & "日期检查.lnk")
sc.TargetPath = path & "日期检查.vbs"
sc.WorkingDirectory = path
 
sc.save

Dim ans : ans = msgbox("已成功添加脚本""日期检查.vbs""到开机启动项。" & VbCrLf &"后续管理请访问任务管理器、控制面板，或各杀毒软件的启动项管理工具。"& VbCrLf &"同时在桌面创建快捷方式？", Vbyesno)
if ans then
	tg = ws.SpecialFolders("Desktop") & "\"
	Set sc = ws.CreateShortcut(tg & "检查日期.lnk")
	sc.TargetPath = path & "日期检查.vbs"
	sc.WorkingDirectory = path
	sc.IconLocation = "%SystemRoot%\System32\SHELL32.dll, 22"
	sc.save
end if