print "Starting test: ChangeSaaSPW"

Environment.Value("basePassword") = "PpmOctane0"
basePassword = Environment.Value("basePassword")
Environment.value("startingPWNum") = 0
startingPWNum = Environment.value("startingPWNum")

'until no more browsers exist
While Browser("creationtime:=0").Exist(0)
'Close the browser
Browser("creationtime:=0").highlight
Browser("creationtime:=0").close
Wend

SystemUtil.Run "Chrome.exe", "https://home.software.microfocus.com/myaccount/?idpId=login.ext.hpe.com#/myProducts"

if Browser("Login").Page("MyAccount - My Products").WebElement("Account").Exist(10) then ' logged in from last time
	RunAction "Logout", oneIteration
	Browser("Login").Navigate "https://home.software.microfocus.com/myaccount/?idpId=login.ext.hpe.com#/myProducts"	
End If

print "About to login with >" & DataTable("userName", dtGlobalSheet) & "<"
Browser("Login").Page("Login").WebEdit("federateLoginName").Set DataTable("userName", dtGlobalSheet) @@ hightlight id_;_Browser("Login").Page("Login").WebEdit("federateLoginName")_;_script infofile_;_ZIP::ssf13.xml_;_
Browser("Login").Page("Login").WebButton("CONTINUE").Click @@ hightlight id_;_Browser("Login").Page("Login").WebButton("CONTINUE")_;_script infofile_;_ZIP::ssf14.xml_;_
Browser("Login").Page("Login").WebEdit("password").Set basePassword & startingPWNum @@ hightlight id_;_Browser("Login").Page("Login").WebEdit("password")_;_script infofile_;_ZIP::ssf15.xml_;_
wait 5
Browser("Login").Page("Login").WebButton("SIGN IN").Click @@ hightlight id_;_Browser("Login").Page("Login").WebButton("SIGN IN")_;_script infofile_;_ZIP::ssf16.xml_;_

if Browser("Login").Page("Login - MyAccount").WebElement("The login name or password").Exist (10) then
	msgbox "Could not login with password: " & basePassword & startingPWNum
	exittest 
ElseIf Browser("Login").Page("Micro Focus - Change Password").WebElement("Your password has expired, enter a new password.").Exist (10) then
	Browser("Login").Page("Micro Focus - Change Password").WebEdit("password").Set basePassword & startingPWNum @@ hightlight id_;_Browser("Login").Page("Micro Focus - Change Password").WebEdit("password")_;_script infofile_;_ZIP::ssf17.xml_;_
	Browser("Login").Page("Micro Focus - Change Password").WebEdit("password1").Set basePassword & startingPWNum + 1 @@ hightlight id_;_Browser("Login").Page("Micro Focus - Change Password").WebEdit("password1")_;_script infofile_;_ZIP::ssf18.xml_;_
	Browser("Login").Page("Micro Focus - Change Password").WebEdit("password2").Set basePassword & startingPWNum + 1 @@ hightlight id_;_Browser("Login").Page("Micro Focus - Change Password").WebEdit("password2")_;_script infofile_;_ZIP::ssf19.xml_;_
	Browser("Login").Page("Micro Focus - Change Password").WebButton("SAVE").Click
end if

