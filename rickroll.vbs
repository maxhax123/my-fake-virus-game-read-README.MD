Dim iURL 
Dim objShell

iURL = "https://www.youtube.com/watch?v=dQw4w9WgXcQ"

set objShell = CreateObject("Shell.Application")
objShell.ShellExecute "chrome.exe", iURL, "", "", 1
