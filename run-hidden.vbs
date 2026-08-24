Set sh = CreateObject("WScript.Shell")
sh.Run """" & WScript.Arguments(0) & """ """ & WScript.Arguments(1) & """", 0, False
