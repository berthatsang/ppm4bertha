'Browser("Search Projects").Page("Search Projects").Link("A/R Billing Upgrade").Click @@ hightlight id_;_Browser("Search Projects").Page("Search Projects").Link("A/R Billing Upgrade")_;_script infofile_;_ZIP::ssf1.xml_;_
Window("Window").Minimize @@ hightlight id_;_21564410_;_script infofile_;_ZIP::ssf27.xml_;_
nRows  = Browser("Search Projects").Page("Search Projects").WebTable("Select Project to View").RowCount
For row = 34 To nRows ' row 1 is the header row so skip
	print row
	staffingProfile = Browser("Search Projects").Page("Search Projects").WebTable("Select Project to View").ChildItemCount (row, 4, "Image") @@ hightlight id_;_Browser("Search Projects").Page("Search Projects").Link("A/R Billing Upgrade")_;_script infofile_;_ZIP::ssf1.xml_;_
	If staffingProfile = 1 Then
        Set projectLink = Browser("Search Projects").Page("Search Projects").WebTable("Select Project to View").ChildItem (row, 2, "Link", 0)
		projectLink.highlight
		projectLink.Click
		Browser("Search Projects").Page("Project Overview").WebElement("staffingTabNav").Click @@ hightlight id_;_Browser("Search Projects").Page("Project Overview").WebElement("staffingTabNav")_;_script infofile_;_ZIP::ssf2.xml_;_
		Browser("Search Projects").Page("Project Overview_2").WebElement("WebElement").Click @@ hightlight id_;_Browser("Search Projects").Page("Project Overview 2").WebElement("WebElement")_;_script infofile_;_ZIP::ssf3.xml_;_
		Browser("Search Projects").Page("Project Overview_2").WebEdit("name").Set "workaround" @@ hightlight id_;_Browser("Search Projects").Page("Project Overview 2").WebEdit("name")_;_script infofile_;_ZIP::ssf4.xml_;_
		Browser("Search Projects").Page("Project Overview_2").WebEdit("role").Set "xoutsource" @@ hightlight id_;_Browser("Search Projects").Page("Project Overview 2").WebEdit("role")_;_script infofile_;_ZIP::ssf5.xml_;_
		
		if Browser("Project Overview").Page("Project Overview").Link("Dismiss PPM Physcal Period Warning").Exist (3) then
			Browser("Project Overview").Page("Project Overview").Link("Dismiss PPM Physcal Period Warning").Click
		end if 

		Browser("Search Projects").Page("Project Overview_2").WebEdit("resourcePool").Set "UI and Web Development - Team 1 (AMS) (FS)" @@ hightlight id_;_Browser("Search Projects").Page("Project Overview 2").WebEdit("resourcePool")_;_script infofile_;_ZIP::ssf6.xml_;_
		Browser("Search Projects").Page("Project Overview_2").WebEdit("costCategory").Set "Employee"

		' got to drag and drop the + sign.  Not easy!!!
		Browser("Search Projects").InsightObject("InsightObject").Click ' got to click in order for the + 'to' appear
		Browser("Search Projects").Page("Project Overview_2").WebElement("spGanttIndicator + sign").highlight
		xloc = Browser("Search Projects").Page("Project Overview_2").WebElement("spGanttIndicator + sign").GetROProperty ("abs_x")
		yloc = Browser("Search Projects").Page("Project Overview_2").WebElement("spGanttIndicator + sign").GetROProperty ("abs_y")
		xwidth = Browser("Search Projects").Page("Project Overview_2").WebElement("spGanttIndicator + sign").GetROProperty ("width")
		yhight = Browser("Search Projects").Page("Project Overview_2").WebElement("spGanttIndicator + sign").GetROProperty ("height")
		
		set DeviceReplay = CreateObject("Mercury.DeviceReplay")
		DeviceReplay.DragAndDrop xloc + xwidth/2, yloc + yhight/2, xloc + 100, yloc + yhight/2, micLeftButn
		Set DeviceReplay = Nothing

		Browser("Search Projects").Page("Project Overview_2").WebEdit("effort").Set "1.0" @@ hightlight id_;_Browser("Search Projects").Page("Project Overview 2").WebEdit("effort")_;_script infofile_;_ZIP::ssf10.xml_;_
		Browser("Search Projects").Page("Project Overview_2").Link("Confirm").Click @@ hightlight id_;_Browser("Search Projects").Page("Project Overview 2").Link("Confirm")_;_script infofile_;_ZIP::ssf11.xml_;_
		Browser("Search Projects").Page("Project Overview_2").Link("Save").Click @@ hightlight id_;_Browser("Search Projects").Page("Project Overview 2").Link("Save")_;_script infofile_;_ZIP::ssf12.xml_;_
		Browser("Search Projects").Page("Project Overview_2").Link("workaround").Check CheckPoint("workaround") @@ script infofile_;_ZIP::ssf14.xml_;_
		
'		Browser("Search Projects").Back
	
		For backToStartd = 1 To 5 Step 1
			
			if Browser("Search Projects").Page("Project Overview_2").Link("Search Projects").Exist (10) then ' standard hack of having to wait after the object appears @@ hightlight id_;_Browser("Search Projects").Page("Project Overview 2").Link("Search Projects")_;_script infofile_;_ZIP::ssf13.xml_;_
				wait 3
				Browser("Search Projects").Page("Project Overview_2").Link("Search Projects").Click @@ hightlight id_;_Browser("Search Projects").Page("Project Overview 2").Link("Search Projects")_;_script infofile_;_ZIP::ssf13.xml_;_
				wait 5
				If Browser("Search Projects").Page("Search Projects").WebTable("Select Project to View").exist (2) Then
					Exit for
				end if
			End If
		next	
	End If
Next


 @@ hightlight id_;_Browser("Search Projects").Page("Project Overview 3").WebElement("pGanttNEW1574706060164")_;_script infofile_;_ZIP::ssf23.xml_;_
foo = 1

