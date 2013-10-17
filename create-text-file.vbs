
dim WshShell

Set WshShell = WScript.CreateObject("WScript.Shell")

WScript.Sleep 100
WshShell.SendKeys "%{Esc}%hw{up}{up}{up}{enter}"
'WScript.Sleep 100
'WshShell.SendKeys "%hw{up}{up}{up}{enter}"
WScript.Sleep 50

