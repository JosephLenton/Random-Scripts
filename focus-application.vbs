
' 
' Gives focus to the application named.
' 
' The intention is that you make a shortcut to this,
' set a keyboard shortcut to that,
' and then you can insta-switch to programs.
' 
' To Use
' ------
' 
' make a shortcut, and point it to:
' 
'       %WINDIR%\System32\wscript.exe <c:\path\to\script.vbs> <application title>
' 
' For example:
' 
'       %WINDIR%\System32\wscript.exe "Z:\git\Random Scripts\focus-application.vbs" "Mozilla FireFox"
' 
'       %WINDIR%\System32\wscript.exe "Z:\git\Random Scripts\focus-application.vbs" "gVim"
' 

dim WshShell, title

title = WScript.Arguments(0)

Set WshShell = WScript.CreateObject("WScript.Shell")

r = WshShell.AppActivate( title )

if r then
    WshShell.SendKeys "% "
    WshShell.SendKeys "x"
else
    WshShell.SendKeys "^{Esc}"
    WScript.Sleep 500
    WshShell.SendKeys title
    WshShell.SendKeys "{Enter}"
end if

