print "Starting test:" & environment("TestName")

' Uncomment this if you want debug statements
foo= 1
	
'Close browsers until no more browsers exist
While Browser("creationtime:=0").Exist(0)
	'Close the browser
	Browser("creationtime:=0").Highlight
	Browser("creationtime:=0").close
	Wait 3
Wend

' Recommend you use chrome - seems to replay faster
SystemUtil.Run "chrome.exe", "http://ppmdemo:8084/itg"


Browser("PPM Logon").Page("PPM Logon").WebEdit("USERNAME").Set "admin" @@ script infofile_;_ZIP::ssf1.xml_;_
Browser("PPM Logon").Page("PPM Logon").WebEdit("PASSWORD").SetSecure "61fc0c118e23c6ef4f56cb64f3e9d036e4925922e68408ff" @@ script infofile_;_ZIP::ssf2.xml_;_
Browser("PPM Logon").Page("PPM Logon").WebElement("label-LOGON_SUBMIT_BUTTON_CAPT").Click @@ script infofile_;_ZIP::ssf3.xml_;_
Browser("PPM Logon").Page("PPM Logon").WebElement("OPEN").Click @@ script infofile_;_ZIP::ssf4.xml_;_
Browser("PPM Logon").Page("PPM Logon").Link("What-if Analysis").Click @@ script infofile_;_ZIP::ssf5.xml_;_

