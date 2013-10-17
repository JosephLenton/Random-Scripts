
dim WshShell

Set WshShell = WScript.CreateObject("WScript.Shell")

WScript.Sleep 100
WshShell.SendKeys "%{Esc}+{F10}"
WScript.Sleep 50

