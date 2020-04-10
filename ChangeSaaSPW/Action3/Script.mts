basePassword = Environment.Value("basePassword")
startingPWNum = Environment.value("startingPWNum")

Browser("MyAccount - My Products").Page("MyAccount - My Products").Link("Settings").Click
' if you don't wait, the next link gives a 404 error!!!!
if not Browser("MyAccount - My Products").Page("MyAccount - My Products").WebElement("authentication email").Exist (10) then
	msgbox "Did not get to Authentication Settings"
end if 

Browser("MyAccount - My Products").Page("MyAccount - My Products").WebElement("My Authentication Settings").Click @@ hightlight id_;_Browser("MyAccount - My Products").Page("MyAccount - My Products").WebElement("My Authentication Settings")_;_script infofile_;_ZIP::ssf2.xml_;_
Browser("MyAccount - My Products").Page("Micro Focus - Authentication").Link("Change").Click @@ hightlight id_;_Browser("MyAccount - My Products").Page("Micro Focus - Authentication").Link("Change")_;_script infofile_;_ZIP::ssf3.xml_;_

For curPassNum = startingPWNum To 12  Step 1	
	print "Setting PW to number " & curPassNum + 1
	Browser("MyAccount - My Products").Page("Micro Focus - Change Password").WebEdit("password").Set basePassword & curPassNum @@ hightlight id_;_Browser("MyAccount - My Products").Page("Micro Focus - Change Password").WebEdit("password")_;_script infofile_;_ZIP::ssf4.xml_;_
	Browser("MyAccount - My Products").Page("Micro Focus - Change Password").WebEdit("password1").Set basePassword & curPassNum + 1 @@ hightlight id_;_Browser("MyAccount - My Products").Page("Micro Focus - Change Password").WebEdit("password1")_;_script infofile_;_ZIP::ssf5.xml_;_
	Browser("MyAccount - My Products").Page("Micro Focus - Change Password").WebEdit("password2").Set basePassword & curPassNum +1 @@ hightlight id_;_Browser("MyAccount - My Products").Page("Micro Focus - Change Password").WebEdit("password2")_;_script infofile_;_ZIP::ssf6.xml_;_
	Browser("MyAccount - My Products").Page("Micro Focus - Change Password").WebButton("SAVE").Click @@ hightlight id_;_Browser("MyAccount - My Products").Page("Micro Focus - Change Password").WebButton("SAVE")_;_script infofile_;_ZIP::ssf7.xml_;_
	If not Browser("MyAccount - My Products").Page("MyAccount - My Products").WebElement("Your new password was created successfully.").Exist (10) then
		msgbox "Password was NOT updated to PpmOctane0" & curPassNum +1
		exittest
	End If
Next

' now back to permanent	
Browser("MyAccount - My Products").Page("Micro Focus - Change Password").WebEdit("password").Set basePassword & curPassNum @@ hightlight id_;_Browser("MyAccount - My Products").Page("Micro Focus - Change Password").WebEdit("password")_;_script infofile_;_ZIP::ssf4.xml_;_
Browser("MyAccount - My Products").Page("Micro Focus - Change Password").WebEdit("password1").Set basePassword & 0 @@ hightlight id_;_Browser("MyAccount - My Products").Page("Micro Focus - Change Password").WebEdit("password1")_;_script infofile_;_ZIP::ssf5.xml_;_
Browser("MyAccount - My Products").Page("Micro Focus - Change Password").WebEdit("password2").Set basePassword & 0 @@ hightlight id_;_Browser("MyAccount - My Products").Page("Micro Focus - Change Password").WebEdit("password2")_;_script infofile_;_ZIP::ssf6.xml_;_
Browser("MyAccount - My Products").Page("Micro Focus - Change Password").WebButton("SAVE").Click
If not Browser("MyAccount - My Products").Page("MyAccount - My Products").WebElement("Your new password was created successfully.").Exist (10) then
		msgbox "Password was NOT updated to PpmOctane00"
		exittest
End If

Browser("MyAccount - My Products").Page("Micro Focus - Change Password_2").Link("Back to settings  >").Click @@ hightlight id_;_Browser("MyAccount - My Products").Page("Micro Focus - Change Password 2").Link("Back to settings  >")_;_script infofile_;_ZIP::ssf12.xml_;_
Browser("MyAccount - My Products").Page("Micro Focus - Authentication_2").Link("Back >").Click @@ hightlight id_;_Browser("MyAccount - My Products").Page("Micro Focus - Authentication 2").Link("Back >")_;_script infofile_;_ZIP::ssf13.xml_;_
Browser("MyAccount - My Products").Page("MyAccount - My Products").Link("PRODUCTS.*").Click @@ hightlight id_;_Browser("MyAccount - My Products").Page("MyAccount - My Products").Link("PRODUCTS 9+")_;_script infofile_;_ZIP::ssf14.xml_;_
