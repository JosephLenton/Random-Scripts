
dim WshShell

Set WshShell = WScript.CreateObject("WScript.Shell")

WScript.Sleep 100
WshShell.SendKeys "%{Esc}+{F10}v"
WScript.Sleep 50

