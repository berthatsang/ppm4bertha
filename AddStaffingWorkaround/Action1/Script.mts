Browser("Project Overview").Navigate "http://ppmdemo:8084/itg/project/SearchProjects.do"
Browser("Project Overview_2").Page("Search Projects").Link("Search").Click @@ script infofile_;_ZIP::ssf38.xml_;_
 @@ script infofile_;_ZIP::ssf38.xml_;_
'Setting.webpackage("ReplayType") = 2 'Mouse
'Setting.webpackage("ReplayType") = 1 'Event

nRows  = Browser("Search Projects").Page("Search Projects").WebTable("Select Project to View").RowCount
For row = 2 To nrows ' row 1 is the header row so skip
	print row & " of " & nRows
	staffingProfile = Browser("Search Projects").Page("Search Projects").WebTable("Select Project to View").ChildItemCount (row, 4, "Image") @@ hightlight id_;_Browser("Search Projects").Page("Search Projects").Link("A/R Billing Upgrade")_;_script infofile_;_ZIP::ssf1.xml_;_
	If staffingProfile = 1 Then
        Set projectLink = Browser("Search Projects").Page("Search Projects").WebTable("Select Project to View").ChildItem (row, 2, "Link", 0)
		projectLink.highlight
		print "project " & projectlink.getroproperty("innertext")	& " has na href of " & projectlink.getroproperty("href")	
		projectLink.DoubleClick
		if not Browser("Search Projects").Page("Project Overview").WebElement("staffingTabNav").Exist(5) then @@ hightlight id_;_Browser("Search Projects").Page("Project Overview").WebElement("staffingTabNav")_;_script infofile_;_ZIP::ssf2.xml_;_
			Setting.webpackage("ReplayType") = 2 'Mouse
			projectLink.highlight
			projectLink.Click
			Setting.webpackage("ReplayType") = 1 'Event
		end if 
		Browser("Search Projects").Page("Project Overview").WebElement("staffingTabNav").Click @@ hightlight id_;_Browser("Search Projects").Page("Project Overview").WebElement("staffingTabNav")_;_script infofile_;_ZIP::ssf2.xml_;_

		If Browser("Project Overview_2").Page("Project Overview").WebElement("gantt view is selected").Exist(5) Then
			' change to table view from gannt
			Browser("Project Overview_2").Page("Project Overview").Link("Change Header").Click
			Browser("Project Overview_2").Page("Change Staffing Profile").WebRadioGroup("editMode").Select "T" @@ script infofile_;_ZIP::ssf32.xml_;_
			Browser("Project Overview_2").Page("Change Staffing Profile").Link("Done").Click	 @@ script infofile_;_ZIP::ssf35.xml_;_
		end if

		Browser("Search Projects").Page("Project Overview_2").WebElement("WebElement").Click @@ hightlight id_;_Browser("Search Projects").Page("Project Overview 2").WebElement("WebElement")_;_script infofile_;_ZIP::ssf3.xml_;_
		Browser("Search Projects").Page("Project Overview_2").WebEdit("name").Set "workaround" @@ hightlight id_;_Browser("Search Projects").Page("Project Overview 2").WebEdit("name")_;_script infofile_;_ZIP::ssf4.xml_;_
		Browser("Search Projects").Page("Project Overview_2").WebEdit("role").Set "xoutsource" @@ hightlight id_;_Browser("Search Projects").Page("Project Overview 2").WebEdit("role")_;_script infofile_;_ZIP::ssf5.xml_;_
		
		if Browser("Project Overview").Page("Project Overview").Link("Dismiss PPM Physcal Period Warning").Exist (3) then
			Browser("Project Overview").Page("Project Overview").Link("Dismiss PPM Physcal Period Warning").Click
		end if 

		Browser("Search Projects").Page("Project Overview_2").WebEdit("resourcePool").Set "UI and Web Development - Team 1 (AMS) (FS)" @@ hightlight id_;_Browser("Search Projects").Page("Project Overview 2").WebEdit("resourcePool")_;_script infofile_;_ZIP::ssf6.xml_;_
		Browser("Search Projects").Page("Project Overview_2").WebEdit("costCategory").Set "Employee"
	
		Browser("Project Overview").Page("Project Overview").WebEdit("Effort").Set "2.00" @@ script infofile_;_ZIP::ssf30.xml_;_
		' Browser("Search Projects").Page("Project Overview_2").Link("Confirm").Click only necessary in gantt, not in table @@ hightlight id_;_Browser("Search Projects").Page("Project Overview 2").Link("Confirm")_;_script infofile_;_ZIP::ssf11.xml_;_
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



foo = 1

