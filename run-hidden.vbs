Set sh = CreateObject("WScript.Shell")
sh.Run "cmd /c " & WScript.Arguments(0), 0, False
