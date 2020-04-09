Browser("Login").Navigate "https://home.software.microfocus.com/myaccount/?idpId=login.ext.hpe.com#/myProducts" @@ hightlight id_;_Browser("Login").Page("Login").WebElement("CONTINUE")_;_script infofile_;_ZIP::ssf1.xml_;_

 @@ hightlight id_;_Browser("Login").Page("Login").WebElement("CONTINUE")_;_script infofile_;_ZIP::ssf5.xml_;_
Browser("Login").Page("Login").WebEdit("federateLoginName").Set "j.banks@microfocus.com" @@ hightlight id_;_Browser("Login").Page("Login").WebEdit("federateLoginName")_;_script infofile_;_ZIP::ssf13.xml_;_
Browser("Login").Page("Login").WebButton("CONTINUE").Click @@ hightlight id_;_Browser("Login").Page("Login").WebButton("CONTINUE")_;_script infofile_;_ZIP::ssf14.xml_;_
Browser("Login").Page("Login").WebEdit("password").Set "PpmOctane00" @@ hightlight id_;_Browser("Login").Page("Login").WebEdit("password")_;_script infofile_;_ZIP::ssf15.xml_;_
Browser("Login").Page("Login").WebButton("SIGN IN").Click @@ hightlight id_;_Browser("Login").Page("Login").WebButton("SIGN IN")_;_script infofile_;_ZIP::ssf16.xml_;_

if Browser("Login").Page("Micro Focus - Change Password").WebElement("Your password has expired, enter a new password.").Exist(10) then
	Browser("Login").Page("Micro Focus - Change Password").WebEdit("password").Set "PpmOctane00" @@ hightlight id_;_Browser("Login").Page("Micro Focus - Change Password").WebEdit("password")_;_script infofile_;_ZIP::ssf17.xml_;_
	Browser("Login").Page("Micro Focus - Change Password").WebEdit("password1").Set "PpmOctane01" @@ hightlight id_;_Browser("Login").Page("Micro Focus - Change Password").WebEdit("password1")_;_script infofile_;_ZIP::ssf18.xml_;_
	Browser("Login").Page("Micro Focus - Change Password").WebEdit("password2").Set "PpmOctane01" @@ hightlight id_;_Browser("Login").Page("Micro Focus - Change Password").WebEdit("password2")_;_script infofile_;_ZIP::ssf19.xml_;_
	Browser("Login").Page("Micro Focus - Change Password").WebButton("SAVE").Click @@ hightlight id_;_Browser("Login").Page("Micro Focus - Change Password").WebButton("SAVE")_;_script infofile_;_ZIP::ssf20.xml_;_
end if 

 @@ hightlight id_;_Browser("Login").Page("Login").WebElement("CONTINUE")_;_script infofile_;_ZIP::ssf12.xml_;_
