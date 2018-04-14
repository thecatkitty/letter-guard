Dim objShell
Set objShell=CreateObject("WScript.Shell")
strCMD="powershell -sta -noProfile -NonInteractive -nologo -file LetterGuard.ps1"
objShell.Run strCMD,0