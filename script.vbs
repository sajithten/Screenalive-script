Dim objResult
Set objShell = Wscript.CreateObject("Wscript.Shell")
Do While True
    objResult = objShell.sendkeys("{NUMLOCK} {NUMLOCK}")
Wscript.Sleep(6000)
Loop
