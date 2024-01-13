Set WshShell = WScript.CreateObject("WScript.Shell")
Set objShell = CreateObject("WScript.Shell")
Set objWshScriptExec = objShell.Exec("cmd /c echo off")
Do
    WshShell.SendKeys "{RANDOM}" ' Simulates pressing a random key
    WshShell.SendKeys "^v" ' Simulates pressing Ctrl+V (paste)
    WshShell.SendKeys "{CLICK}" ' Simulates a mouse click
    WScript.Sleep 100 ' Waits for 100 milliseconds
Loop
