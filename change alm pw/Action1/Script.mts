print "Starting test: " & Environment("TestName")

'until no more browsers exist
While Browser("creationtime:=0").Exist(0)
	'Close the browser
	Browser("creationtime:=0").highlight
	Browser("creationtime:=0").close
	Wait 5
Wend

SystemUtil.Run "iexplore.exe", "http://nimbusserver.mfadvantageinc.com:8082/qcbin/"

Browser("Micro Focus Application").Page("Site Admin Page").Link("Site Administration").Click @@ hightlight id_;_Browser("Micro Focus Application 2").Page("Micro Focus Application").Link("Site Administration")_;_script infofile_;_ZIP::ssf50.xml_;_
Browser("Micro Focus Application").Page("Site Admin Page").DelphiObject("Username").Type "admin" @@ hightlight id_;_4394704_;_script infofile_;_ZIP::ssf51.xml_;_
'Browser("Micro Focus Application").Page("Micro Focus Application").DelphiObject("DelphiObject").Type  micTab  @@ hightlight id_;_4394704_;_script infofile_;_ZIP::ssf52.xml_;_
Browser("Micro Focus Application").Page("Site Admin Page").DelphiObject("Password").Type "Password1" @@ hightlight id_;_1052342_;_script infofile_;_ZIP::ssf53.xml_;_
Browser("Micro Focus Application").Page("Site Admin Page").DelphiObject("LoginButton").Click '51,17 @@ hightlight id_;_1380024_;_script infofile_;_ZIP::ssf54.xml_;_

If Browser("Micro Focus Application").InsightObject("Site Users Tab").Exist (20) THEN @@ hightlight id_;_6_;_script infofile_;_ZIP::ssf56.xml_;_
	Wait 3
End If
Browser("Micro Focus Application").InsightObject("Site Users Tab").Click

Browser("Micro Focus Application").Page("Site Admin Page").DelphiObject("Find Name Textbox").highlight
Browser("Micro Focus Application").Page("Site Admin Page").DelphiObject("Find Name Textbox").DblClick 5,5
Browser("Micro Focus Application").Page("Site Admin Page").DelphiObject("Find Name Textbox").Type  micBack @@ hightlight id_;_396958_;_script infofile_;_ZIP::ssf12.xml_;_
'Window("Internet Explorer").WinObject("std_attached_and_text_tag_quer").Type DataTable("UserName", dtGlobalSheet) @@ hightlight id_;_396958_;_script infofile_;_ZIP::ssf2.xml_;_
'Window("Internet Explorer").WinObject("std_attached_and_text_tag_quer").Type  micReturn @@ hightlight id_;_396958_;_script infofile_;_ZIP::ssf3.xml_;_

Browser("Micro Focus Application").Page("Site Admin Page").DelphiObject("Find Name Textbox").Type  DataTable("UserName", dtGlobalSheet)
Browser("Micro Focus Application").Page("Site Admin Page").DelphiObject("Find Name Textbox").Type  micReturn

Browser("Micro Focus Application").InsightObject("Password Button").Click @@ hightlight id_;_13_;_script infofile_;_ZIP::ssf4.xml_;_

DelphiWindow("Set User Password Dialog").DelphiObject("New Password Textbox").Click ' 62,5 @@ hightlight id_;_921212_;_script infofile_;_ZIP::ssf58.xml_;_
DelphiWindow("Set User Password Dialog").DelphiObject("New Password Textbox").Type "mfDemo$20" @@ hightlight id_;_921212_;_script infofile_;_ZIP::ssf59.xml_;_
DelphiWindow("Set User Password Dialog").DelphiObject("New Password Textbox").Type  micTab  @@ hightlight id_;_921212_;_script infofile_;_ZIP::ssf60.xml_;_
DelphiWindow("Set User Password Dialog").DelphiObject("Confirm Password Textbox").Type "mfDemo$20" @@ hightlight id_;_593076_;_script infofile_;_ZIP::ssf61.xml_;_
'DelphiWindow("Set User Password Dialog").DelphiObject("Confirm Password Textbox").Type  micReturn  @@ hightlight id_;_593076_;_script infofile_;_ZIP::ssf62.xml_;_
DelphiWindow("Set User Password Dialog").DelphiObject("OK button").Click  
DelphiWindow("Password Changed Dialog").DelphiObject("OK button").Click '39,9 @@ hightlight id_;_1115920_;_script infofile_;_ZIP::ssf63.xml_;_


'DelphiWindow("Name").InsightObject("InsightObject").Click
'Window("Window").Window("RegExpWndTitle").WinObject("PW").Type "mfDemo$20" @@ hightlight id_;_855240_;_script infofile_;_ZIP::ssf6.xml_;_
'DelphiWindow("Name").InsightObject("InsightObject_2").Click
'Window("Window").Window("RegExpWndTitle").WinObject("PW Confirm").Type "mfDemo$20" @@ hightlight id_;_1051460_;_script infofile_;_ZIP::ssf8.xml_;_
'DelphiWindow("Name").DelphiObject("OK in PW set").Click

'Browser("Micro Focus Application").InsightObject("OK Button").Click ' ok in confirm success @@ hightlight id_;_2_;_script infofile_;_ZIP::ssf13.xml_;_


''Until no more ALM/QCs exist
'While Window("Site Administration http://nim").Exist(0)
'	'Close ALM/QC
'	Window("Site Administration http://nim").highlight
'	Window("Site Administration http://nim").close
'Wend
'SystemUtil.Run "ALMClientLauncher.exe", "", "C:\Program Files\Micro Focus\ALMClient"

'Window("ALM Client Launcher").WinObject("WindowsForms10.EDIT.app.0.34f5").Type "http://nimbusserver.mfadvantageinc.com:8082/qcbin/siteadmin" @@ hightlight id_;_3673808_;_script infofile_;_ZIP::ssf31.xml_;_
'Window("ALM Client Launcher").WinObject("WindowsForms10.EDIT.app.0.34f5").Type  micReturn  @@ hightlight id_;_1444128_;_script infofile_;_ZIP::ssf38.xml_;_
'Window("Site Administration http://nim").DelphiObject("DelphiObject_3").Type "admin" @@ hightlight id_;_4129838_;_script infofile_;_ZIP::ssf39.xml_;_
'Window("Site Administration http://nim").DelphiObject("DelphiObject_3").Type  micTab  @@ hightlight id_;_4129838_;_script infofile_;_ZIP::ssf40.xml_;_
'Window("Site Administration http://nim").DelphiObject("DelphiObject").Type "Password1"
'Window("Site Administration http://nim").DelphiObject("DelphiObject").Type  micTab  @@ hightlight id_;_1377324_;_script infofile_;_ZIP::ssf41.xml_;_
'Window("Site Administration http://nim").DelphiObject("DelphiObject").Type  micReturn 
