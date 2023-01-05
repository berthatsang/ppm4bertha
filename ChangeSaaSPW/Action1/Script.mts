' Test 05/01/2023
print "Starting test: " & Environment("TestName")

'  SaaS doesn't allow you to use the same password within the past 12 passwords.  
'  This script logs in as a user and changes the password 12 times and finally sets 
'  the password to PpmOctane00.
'  The initial password used to login with, must end in "00" e.g. R0l3sP14y00.  
' It must be entered into the variable basePassword below with the trailing 0 
' removed, prior to this script being run.

'Environment.Value("basePassword") = "PpmOctane0"
Environment.Value("basePassword") = "R0l3sP14y0"
basePassword = Environment.Value("basePassword")

Environment.Value("startingPWNum") = 0
startingPWNum = Environment.Value("startingPWNum")

'until no more browsers exist
While Browser("creationtime:=0").Exist(0)
	'Close the browser
	Browser("creationtime:=0").highlight
	Browser("creationtime:=0").close
Wend

SystemUtil.Run "Chrome.exe", "https://home.software.microfocus.com/myaccount/?idpId=login.ext.hpe.com#/myProducts"

If Browser("Login").Page("MyAccount - My Products").WebElement("Account").Exist(10) Then ' logged in from last time
	RunAction "Logout", oneIteration
	Browser("Login").Navigate "https://home.software.microfocus.com/myaccount/?idpId=login.ext.hpe.com#/myProducts"	
End If

curPassword=basePassword & startingPWNum
curPassword=basePassword & startingPWNum
print "About to login with: " & DataTable("userName", dtGlobalSheet) & " and curPassword: " & curPassword 

Browser("Login").Page("Login").WebEdit("federateLoginName").Set DataTable("userName", dtGlobalSheet) @@ hightlight id_;_Browser("Login").Page("Login").WebEdit("federateLoginName")_;_script infofile_;_ZIP::ssf13.xml_;_
Browser("Login").Page("Login").WebButton("CONTINUE").Click @@ hightlight id_;_Browser("Login").Page("Login").WebButton("CONTINUE")_;_script infofile_;_ZIP::ssf14.xml_;_
Browser("Login").Page("Login").WebEdit("password").Set curPassword 

Browser("Login").Page("Login").WebButton("SIGN IN").Click @@ hightlight id_;_Browser("Login").Page("Login").WebButton("SIGN IN")_;_script infofile_;_ZIP::ssf16.xml_;_
Browser("Login").Page("Login").WebButton("SIGN IN").Click @@ hightlight id_;_Browser("Login").Page("Login").WebButton("SIGN IN")_;_script infofile_;_ZIP::ssf16.xml_;_

If Browser("Login").Page("Login - MyAccount").WebElement("The login name or password").Exist (10) Then
	MsgBox "Could not login with password: " & curPassword 
	print "Could not login with password: " & curPassword 
	exittest 
	'Browser("Login").Page("Micro Focus - Change Password").WebEdit("password").Set basePassword & startingPWNum @@ hightlight id_;_Browser("Login").Page("Micro Focus - Change Password").WebEdit("password")_;_script infofile_;_ZIP::ssf17.xml_;_
	'Browser("Login").Page("Micro Focus - Change Password").WebEdit("password1").Set basePassword & startingPWNum + 1 @@ hightlight id_;_Browser("Login").Page("Micro Focus - Change Password").WebEdit("password1")_;_script infofile_;_ZIP::ssf18.xml_;_
	'Browser("Login").Page("Micro Focus - Change Password").WebEdit("password2").Set basePassword & startingPWNum + 1 @@ hightlight id_;_Browser("Login").Page("Micro Focus - Change Password").WebEdit("password2")_;_script infofile_;_ZIP::ssf19.xml_;_

	newPassword = curPassword & startingPWNum + 1
	print "newPassword: " & newPassword
	Browser("Login").Page("Micro Focus - Change Password").WebEdit("password").Set curPassword
	Browser("Login").Page("Micro Focus - Change Password").WebEdit("password1").Set newPassword
	Browser("Login").Page("Micro Focus - Change Password").WebEdit("password2").Set newPassword
	
	Browser("Login").Page("Micro Focus - Change Password").WebButton("SAVE").Click
End If

