basePassword = Environment.Value("basePassword")
startingPWNum = Environment.value("startingPWNum")

Browser("MyAccount - My Products").Page("MyAccount - My Products_2").WebElement("WebElement").Click @@ script infofile_;_ZIP::ssf16.xml_;_

' if you don't wait, the next link gives a 404 error!!!!
If Not Browser("MyAccount - My Products").Page("MyAccount - My Products").WebElement("authentication email").Exist (10) Then
	MsgBox "Did not get to Authentication Settings"
End If 

Browser("MyAccount - My Products").Page("MyAccount - My Products_2").WebElement("My Authentication Settings").Click @@ script infofile_;_ZIP::ssf17.xml_;_
Browser("MyAccount - My Products").Page("Micro Focus - Authentication_3").Link("Change").Click @@ script infofile_;_ZIP::ssf18.xml_;_

For curPassNum = startingPWNum To 12  Step 1	
	print "Setting PW to number " & curPassNum + 1
	curPassword=basePassword & curPassNum
	newPassword=basePassword & curPassNum + 1
	print "Changing password from: " & curPassword & " to: " & newPassword 
	
	Browser("MyAccount - My Products").Page("Micro Focus - Change Password").WebEdit("password").Set curPassword @@ hightlight id_;_Browser("MyAccount - My Products").Page("Micro Focus - Change Password").WebEdit("password")_;_script infofile_;_ZIP::ssf4.xml_;_
	Browser("MyAccount - My Products").Page("Micro Focus - Change Password").WebEdit("password1").Set newPassword 
	Browser("MyAccount - My Products").Page("Micro Focus - Change Password").WebEdit("password2").Set newPassword
	
	Browser("MyAccount - My Products").Page("Micro Focus - Change Password").WebButton("SAVE").Click @@ hightlight id_;_Browser("MyAccount - My Products").Page("Micro Focus - Change Password").WebButton("SAVE")_;_script infofile_;_ZIP::ssf7.xml_;_
	If Not Browser("MyAccount - My Products").Page("MyAccount - My Products").WebElement("Your new password was created successfully.").Exist (10) then
		' MsgBox "Password was NOT updated to PpmOctane0" & curPassNum +1
		MsgBox "Password was NOT updated to: " & newPassword
		Print "Password was NOT updated to: " & newPassword
		exittest
	End If
Next

' Now change it to the final password which is always PpmOctane00
finalPassword="PpmOctane00"
print "Changing password from: " & newPassword & " to: " & finalPassword 
Browser("MyAccount - My Products").Page("Micro Focus - Change Password").WebEdit("password").Set newPassword 
Browser("MyAccount - My Products").Page("Micro Focus - Change Password").WebEdit("password1").Set finalPassword @@ hightlight id_;_Browser("MyAccount - My Products").Page("Micro Focus - Change Password").WebEdit("password2")_;_script infofile_;_ZIP::ssf6.xml_;_
Browser("MyAccount - My Products").Page("Micro Focus - Change Password").WebEdit("password2").Set finalPassword @@ hightlight id_;_Browser("MyAccount - My Products").Page("Micro Focus - Change Password").WebEdit("password2")_;_script infofile_;_ZIP::ssf6.xml_;_

Browser("MyAccount - My Products").Page("Micro Focus - Change Password").WebButton("SAVE").Click
If Not Browser("MyAccount - My Products").Page("MyAccount - My Products").WebElement("Your new password was created successfully.").Exist (10) then
	MsgBox "Password was NOT updated to " & finalPassword
	print "Password was NOT updated to " & finalPassword
	exittest
End If

Browser("MyAccount - My Products").Page("Micro Focus - Change Password_2").Link("Back to settings  >").Click @@ hightlight id_;_Browser("MyAccount - My Products").Page("Micro Focus - Change Password 2").Link("Back to settings  >")_;_script infofile_;_ZIP::ssf12.xml_;_
Browser("MyAccount - My Products").Page("Micro Focus - Authentication_2").Link("Back >").Click @@ hightlight id_;_Browser("MyAccount - My Products").Page("Micro Focus - Authentication 2").Link("Back >")_;_script infofile_;_ZIP::ssf13.xml_;_

'Not sure why, but clicking back seems to log you out!  Great for me on this run as I only wanted to change one user, but more problematic if you want to do all users.
' Maybe next time, don't bother going back nicely and go straight to PRODUCTS tab.
Browser("MyAccount - My Products").Page("MyAccount - My Products").Link("PRODUCTS.*").Click @@ hightlight id_;_Browser("MyAccount - My Products").Page("MyAccount - My Products").Link("PRODUCTS 9+")_;_script infofile_;_ZIP::ssf14.xml_;_
