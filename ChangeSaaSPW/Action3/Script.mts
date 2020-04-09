Browser("MyAccount - My Products").Page("MyAccount - My Products").Link("Settings").Click @@ hightlight id_;_Browser("MyAccount - My Products").Page("MyAccount - My Products").Link("Settings")_;_script infofile_;_ZIP::ssf1.xml_;_
wait 10 ' if you don't wait, the next link gives a 404 error!!!!
Browser("MyAccount - My Products").Page("MyAccount - My Products").WebElement("My Authentication Settings").Click @@ hightlight id_;_Browser("MyAccount - My Products").Page("MyAccount - My Products").WebElement("My Authentication Settings")_;_script infofile_;_ZIP::ssf2.xml_;_
Browser("MyAccount - My Products").Page("Micro Focus - Authentication").Link("Change").Click @@ hightlight id_;_Browser("MyAccount - My Products").Page("Micro Focus - Authentication").Link("Change")_;_script infofile_;_ZIP::ssf3.xml_;_

Const InitialPWNum = 1

For curPassNum = InitialPWNum To 12  Step 1	
	print "Setting PW to number " & curPassNum + 1
	Browser("MyAccount - My Products").Page("Micro Focus - Change Password").WebEdit("password").Set "PpmOctane0" & curPassNum @@ hightlight id_;_Browser("MyAccount - My Products").Page("Micro Focus - Change Password").WebEdit("password")_;_script infofile_;_ZIP::ssf4.xml_;_
	Browser("MyAccount - My Products").Page("Micro Focus - Change Password").WebEdit("password1").Set "PpmOctane0" & curPassNum + 1 @@ hightlight id_;_Browser("MyAccount - My Products").Page("Micro Focus - Change Password").WebEdit("password1")_;_script infofile_;_ZIP::ssf5.xml_;_
	Browser("MyAccount - My Products").Page("Micro Focus - Change Password").WebEdit("password2").Set "PpmOctane0" & curPassNum +1 @@ hightlight id_;_Browser("MyAccount - My Products").Page("Micro Focus - Change Password").WebEdit("password2")_;_script infofile_;_ZIP::ssf6.xml_;_
	Browser("MyAccount - My Products").Page("Micro Focus - Change Password").WebButton("SAVE").Click @@ hightlight id_;_Browser("MyAccount - My Products").Page("Micro Focus - Change Password").WebButton("SAVE")_;_script infofile_;_ZIP::ssf7.xml_;_
	If not Browser("MyAccount - My Products").Page("MyAccount - My Products").WebElement("Your new password was created successfully.").Exist (10) then
		msgbox "Password was NOT updated to PpmOctane0" & curPassNum +1
	End If
Next

' now back to permanent	
Browser("MyAccount - My Products").Page("Micro Focus - Change Password").WebEdit("password").Set "PpmOctane0" & curPassNum @@ hightlight id_;_Browser("MyAccount - My Products").Page("Micro Focus - Change Password").WebEdit("password")_;_script infofile_;_ZIP::ssf4.xml_;_
Browser("MyAccount - My Products").Page("Micro Focus - Change Password").WebEdit("password1").Set "PpmOctane00" @@ hightlight id_;_Browser("MyAccount - My Products").Page("Micro Focus - Change Password").WebEdit("password1")_;_script infofile_;_ZIP::ssf5.xml_;_
Browser("MyAccount - My Products").Page("Micro Focus - Change Password").WebEdit("password2").Set "PpmOctane00" @@ hightlight id_;_Browser("MyAccount - My Products").Page("Micro Focus - Change Password").WebEdit("password2")_;_script infofile_;_ZIP::ssf6.xml_;_
Browser("MyAccount - My Products").Page("Micro Focus - Change Password").WebButton("SAVE").Click
If not Browser("MyAccount - My Products").Page("MyAccount - My Products").WebElement("Your new password was created successfully.").Exist (10) then
		msgbox "Password was NOT updated to PpmOctane00"
End If

Browser("MyAccount - My Products").Page("Micro Focus - Change Password_2").Link("Back to settings  >").Click @@ hightlight id_;_Browser("MyAccount - My Products").Page("Micro Focus - Change Password 2").Link("Back to settings  >")_;_script infofile_;_ZIP::ssf12.xml_;_
Browser("MyAccount - My Products").Page("Micro Focus - Authentication_2").Link("Back >").Click @@ hightlight id_;_Browser("MyAccount - My Products").Page("Micro Focus - Authentication 2").Link("Back >")_;_script infofile_;_ZIP::ssf13.xml_;_
Browser("MyAccount - My Products").Page("MyAccount - My Products").Link("PRODUCTS 9+").Click @@ hightlight id_;_Browser("MyAccount - My Products").Page("MyAccount - My Products").Link("PRODUCTS 9+")_;_script infofile_;_ZIP::ssf14.xml_;_
