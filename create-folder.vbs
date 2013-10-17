
dim WshShell

Set WshShell = WScript.CreateObject("WScript.Shell")

WScript.Sleep 100
WshShell.SendKeys "%{Esc}%hn"
WScript.Sleep 50

