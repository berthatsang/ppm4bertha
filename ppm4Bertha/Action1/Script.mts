'Browser("Search Projects").Page("Search Projects").Link("A/R Billing Upgrade").Click @@ hightlight id_;_Browser("Search Projects").Page("Search Projects").Link("A/R Billing Upgrade")_;_script infofile_;_ZIP::ssf1.xml_;_
nRows  = Browser("Search Projects").Page("Search Projects").WebTable("Select Project to View").RowCount
For row = 3 To nRows -6 ' row 1 is the header row so skip
	staffingProfile = Browser("Search Projects").Page("Search Projects").WebTable("Select Project to View").ChildItemCount (row, 4, "Image") @@ hightlight id_;_Browser("Search Projects").Page("Search Projects").Link("A/R Billing Upgrade")_;_script infofile_;_ZIP::ssf1.xml_;_
	If staffingProfile = 1 Then
        Set projectLink = Browser("Search Projects").Page("Search Projects").WebTable("Select Project to View").ChildItem (row, 2, "Link", 0)
		projectLink.highlight
		projectLink.Click
		Browser("Search Projects").Page("Project Overview").WebElement("staffingTabNav").Click @@ hightlight id_;_Browser("Search Projects").Page("Project Overview").WebElement("staffingTabNav")_;_script infofile_;_ZIP::ssf2.xml_;_
		Browser("Search Projects").Page("Project Overview_2").WebElement("WebElement").Click @@ hightlight id_;_Browser("Search Projects").Page("Project Overview 2").WebElement("WebElement")_;_script infofile_;_ZIP::ssf3.xml_;_
		Browser("Search Projects").Page("Project Overview_2").WebEdit("name").Set "workaround" @@ hightlight id_;_Browser("Search Projects").Page("Project Overview 2").WebEdit("name")_;_script infofile_;_ZIP::ssf4.xml_;_
		Browser("Search Projects").Page("Project Overview_2").WebEdit("role").Set "xoutsource" @@ hightlight id_;_Browser("Search Projects").Page("Project Overview 2").WebEdit("role")_;_script infofile_;_ZIP::ssf5.xml_;_
		Browser("Search Projects").Page("Project Overview_2").WebEdit("resourcePool").Set "UI and Web Development - Team 1 (AMS) (FS)" @@ hightlight id_;_Browser("Search Projects").Page("Project Overview 2").WebEdit("resourcePool")_;_script infofile_;_ZIP::ssf6.xml_;_
		'Browser("Search Projects").Page("Project Overview_2").Frame("availFrame").WebElement("WebTable").Click @@ hightlight id_;_Browser("Search Projects").Page("Project Overview 2").Frame("availFrame").WebElement("WebTable")_;_script infofile_;_ZIP::ssf7.xml_;_
		Browser("Search Projects").Page("Project Overview_2").WebEdit("costCategory").Set "employee" @@ hightlight id_;_Browser("Search Projects").Page("Project Overview 2").WebEdit("costCategory")_;_script infofile_;_ZIP::ssf8.xml_;_
		Browser("Search Projects").Page("Project Overview_2").WebElement("Total Positions:  1 Fulfilled").Click @@ hightlight id_;_Browser("Search Projects").Page("Project Overview 2").WebElement("Total Positions:  1 Fulfilled")_;_script infofile_;_ZIP::ssf9.xml_;_
		Browser("Search Projects").Page("Project Overview_2").WebEdit("effort").Set "1.0" @@ hightlight id_;_Browser("Search Projects").Page("Project Overview 2").WebEdit("effort")_;_script infofile_;_ZIP::ssf10.xml_;_
		Browser("Search Projects").Page("Project Overview_2").Link("Confirm").Click @@ hightlight id_;_Browser("Search Projects").Page("Project Overview 2").Link("Confirm")_;_script infofile_;_ZIP::ssf11.xml_;_
		Browser("Search Projects").Page("Project Overview_2").Link("Save").Click @@ hightlight id_;_Browser("Search Projects").Page("Project Overview 2").Link("Save")_;_script infofile_;_ZIP::ssf12.xml_;_
		Browser("Search Projects").Page("Project Overview_2").Link("workaround").Check CheckPoint("workaround") @@ script infofile_;_ZIP::ssf14.xml_;_
		Browser("Search Projects").Page("Project Overview_2").Link("Search Projects").Click @@ hightlight id_;_Browser("Search Projects").Page("Project Overview 2").Link("Search Projects")_;_script infofile_;_ZIP::ssf13.xml_;_
		
	End If
	print staffingProfile
Next

