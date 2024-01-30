CreateObject("wscript.shell").run "E:\Ipconfig.txt"

' Create object for sending keys
Set x = CreateObject("WScript.shell")

' Open start menu
WScript.sleep(500)
x.Sendkeys "^{ESC}"
WScript.sleep(2000)

' Open command promt
x.Sendkeys "cmd"
WScript.sleep(500)
x.Sendkeys "{ENTER}"
WScript.sleep(1800)

' Enter the ipconfig command and copy results
x.Sendkeys "ipconfig | clip"
WScript.sleep(100)
x.Sendkeys "{ENTER}"
WScript.sleep(100)

' exit command prompt
x.sendkeys "exit"
WScript.sleep(100)
x.sendkeys "{ENTER}"

' open "Results.txt" and paste results
CreateObject("wscript.shell").run "E:\Results.txt"
Wscript.sleep(1000)
x.sendkeys "^v"
WScript.sleep(500)

' Save results and exit
x.sendkeys "^s"
WScript.sleep(500)
x.sendkeys "%{F4}"