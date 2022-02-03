print "Starting test:" & environment("TestName")

'close browsers until no more browsers exist
While Browser("creationtime:=0").Exist(0)
'Close the browser
Browser("creationtime:=0").highlight
Browser("creationtime:=0").close
wait 3
Wend

' recommend you use chrome - seems to replay faster
SystemUtil.Run "chrome.exe", "http://ppmdemo:8084/itg"

Browser("PPM Logon").Page("PPM Logon").WebEdit("USERNAME").Set DataTable("p_username", dtGlobalSheet) @@ hightlight id_;_Browser("PPM Logon").Page("PPM Logon").WebEdit("USERNAME")_;_script infofile_;_ZIP::ssf1.xml_;_
Browser("PPM Logon").Page("PPM Logon").WebEdit("PASSWORD").SetSecure "61e6fa48736cdc3b059850e9ad60d16168f5b385f5add20c" @@ hightlight id_;_Browser("PPM Logon").Page("PPM Logon").WebEdit("PASSWORD")_;_script infofile_;_ZIP::ssf2.xml_;_
Browser("PPM Logon").Page("PPM Logon").WebButton("Submit Query").Click @@ hightlight id_;_Browser("PPM Logon").Page("PPM Logon").WebButton("Submit Query")_;_script infofile_;_ZIP::ssf3.xml_;_
Browser("PPM Logon").Page("PPM Logon").Link("OPEN").Click @@ hightlight id_;_Browser("PPM Logon").Page("PPM Logon").Link("OPEN")_;_script infofile_;_ZIP::ssf4.xml_;_
Browser("PPM Logon").Page("PPM Logon").Link("What-if Analysis").Click @@ hightlight id_;_Browser("PPM Logon").Page("PPM Logon").Link("What-if Analysis")_;_script infofile_;_ZIP::ssf5.xml_;_

' Set zoom to 100%
Window("Google Chrome").Type micCtrlDwn + "0" + micCtrlUp
