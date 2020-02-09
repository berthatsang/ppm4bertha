Browser("Project Overview_2").Page("Staffing Profile").Link("Search Staffing Profile").Click

Browser("Project Overview").Navigate "http://ppmdemo:8084/itg/project/SearchProjects.do"
Browser("Project Overview_2").Page("Search Projects").Link("SEARCH_2").Click @@ script infofile_;_ZIP::ssf39.xml_;_
Browser("Project Overview_2").Page("Search Projects").Link("Staffing Profiles").Click @@ script infofile_;_ZIP::ssf40.xml_;_
Browser("Project Overview_2").Page("Search Staffing Profile").Image("Choose Status").Click @@ script infofile_;_ZIP::ssf41.xml_;_

' add all status except baseline
Browser("Project Overview_2").Page("Search Staffing Profile").Frame("availFrame").WebElement("WebTable_2").Click @@ hightlight id_;_Browser("Project Overview 2").Page("Search Staffing Profile").Frame("availFrame").WebElement("WebTable 2")_;_script infofile_;_ZIP::ssf56.xml_;_
Browser("Project Overview_2").Page("Search Staffing Profile").Frame("availFrame").WebElement("WebTable_3").Click @@ hightlight id_;_Browser("Project Overview 2").Page("Search Staffing Profile").Frame("availFrame").WebElement("WebTable 3")_;_script infofile_;_ZIP::ssf57.xml_;_
Browser("Project Overview_2").Page("Search Staffing Profile").Frame("availFrame").WebElement("WebTable_4").Click @@ hightlight id_;_Browser("Project Overview 2").Page("Search Staffing Profile").Frame("availFrame").WebElement("WebTable 4")_;_script infofile_;_ZIP::ssf58.xml_;_
Browser("Project Overview_2").Page("Search Staffing Profile").Frame("availFrame").WebElement("WebTable").Click @@ hightlight id_;_Browser("Project Overview 2").Page("Search Staffing Profile").Frame("availFrame").WebElement("WebTable")_;_script infofile_;_ZIP::ssf59.xml_;_
Browser("Project Overview_2").Page("Search Staffing Profile").Frame("availFrame").WebElement("WebTable_5").Click @@ hightlight id_;_Browser("Project Overview 2").Page("Search Staffing Profile").Frame("availFrame").WebElement("WebTable 5")_;_script infofile_;_ZIP::ssf60.xml_;_
Browser("Project Overview_2").Page("Search Staffing Profile").Frame("availFrame").WebElement("WebTable_6").Click @@ hightlight id_;_Browser("Project Overview 2").Page("Search Staffing Profile").Frame("availFrame").WebElement("WebTable 6")_;_script infofile_;_ZIP::ssf61.xml_;_
Browser("Project Overview_2").Page("Search Staffing Profile").Frame("autoCompleteDialogIF").WebButton("btnOK").Click @@ script infofile_;_ZIP::ssf46.xml_;_
Browser("Project Overview_2").Page("Search Staffing Profile").WebEdit("pageSize").Set "500" @@ script infofile_;_ZIP::ssf47.xml_;_
Browser("Project Overview_2").Page("Search Staffing Profile").WebButton("PushButton_1581043286106_2773").Click @@ script infofile_;_ZIP::ssf48.xml_;_
 @@ script infofile_;_ZIP::ssf38.xml_;_
'Setting.webpackage("ReplayType") = 2 'Mouse
'Setting.webpackage("ReplayType") = 1 'Event

'Browser("Project Overview_2").Page("Search Staffing Profile").WebTable("Profile Name").Check CheckPoint("Profile Name") @@ hightlight id_;_Browser("Project Overview 2").Page("Search Staffing Profile").WebTable("Profile Name")_;_script infofile_;_ZIP::ssf62.xml_;_


nRows  = Browser("Project Overview_2").Page("Search Staffing Profile").WebTable("Profile Name").RowCount
For row = 2 To nrows ' row 1 is the header row so skip
	print row & " of " & nRows
	set projectLink = Browser("Search Projects").Page("Search Projects").WebTable("Select Project to View").ChildItem (row, 1, "Link",0)
	projectLink.highlight
	projName = projectlink.getroproperty("innertext")
	projHref = projectlink.getroproperty("href")
	projectLink.DoubleClick
	If projName = "MSP Integration" Then
		print "skipping project " & projName & " because it is connected to MS Project"
	ElseIf  Browser("Project Overview_2").Page("Staffing Profile").Link("workaround").Exist (10) then ' already exists, we are done @@ script infofile_;_ZIP::ssf64.xml_;_
	' do nothing
		print "project " & projName	& " already has workaround"
		Browser("Project Overview_2").Page("Staffing Profile").Link("Search Staffing Profile").Click
	else
		print "adding workaround to project " & projName & " which has an href of " & projHref 
	
		If Browser("Project Overview_2").Page("Project Overview").WebElement("gantt view is selected").Exist(5) Then
			' change to table view from gannt
			Browser("Project Overview_2").Page("Project Overview").Link("Change Header").Click
			Browser("Project Overview_2").Page("Change Staffing Profile").WebRadioGroup("editMode").Select "T" @@ script infofile_;_ZIP::ssf32.xml_;_
			Browser("Project Overview_2").Page("Change Staffing Profile").Link("Done").Click @@ script infofile_;_ZIP::ssf35.xml_;_
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

		
	end if



Next



foo = 1

