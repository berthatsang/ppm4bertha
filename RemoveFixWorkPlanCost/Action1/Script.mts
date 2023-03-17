print "Starting test: " & Environment("TestName")

'until no more browsers exist
While Browser("creationtime:=0").Exist(0)
	'Close the browser
	Browser("creationtime:=0").highlight
	Browser("creationtime:=0").close
Wend

SystemUtil.Run "Chrome.exe", "http://ppmdemo:8084/itg/project/SearchProjects.do"

Browser("View Work Plan_2").Page("PPM Logon").WebEdit("USERNAME").Set "admin" @@ script infofile_;_ZIP::ssf176.xml_;_
Browser("View Work Plan_2").Page("PPM Logon").WebEdit("PASSWORD").SetSecure "61b72773fc97618e2603ae2e84b37b0997032e65c4d6d405" @@ script infofile_;_ZIP::ssf177.xml_;_
Browser("View Work Plan_2").Page("PPM Logon").WebElement("label-LOGON_SUBMIT_BUTTON_CAPT").Click @@ script infofile_;_ZIP::ssf178.xml_;_

' Set the Results Displayed Per Page textbox to > 50 so that ALL projects are displayed
Browser("Search Projects_2").Page("Search Projects").WebEdit("pageSize").Set "550" ' ridiculously large number to get them all @@ script infofile_;_ZIP::ssf101.xml_;_
'Browser("Search Projects_2").Page("Search Projects").Link("Search").Click @@ script infofile_;_ZIP::ssf102.xml_;_

'Browser("View Work Plan_2").Page("PPM Logon").WebButton("SEARCH_BUTTON_LINK").Click @@ script infofile_;_ZIP::ssf179.xml_;_
'Browser("View Work Plan_2").Page("Search Projects").WebButton("SEARCH_BUTTON_LINK").Click
wait 5
Browser("View Work Plan_2").Page("PPM Logon").WebButton("SEARCH_BUTTON_LINK").Highlight
Browser("View Work Plan_2").Page("PPM Logon").WebButton("SEARCH_BUTTON_LINK").Click

' From Jaya
'Browser("View Work Plan_2").Page("PPM Logon").WebButton("value:=Search,xpath:=//TR/TD/TABLE/TBODY/TR/TD/A[@role='button' and normalize-space()='Search']").Click
'Browser("View Work Plan_2").Page("PPM Logon").WebButton("name:=Search").Click

nRows  = Browser("Search Projects").Page("Search Projects").WebTable("Select Project to View").RowCount
For row = 2 To nRows ' row 1 is the header row so skip
	projectName = Browser("Search Projects").Page("Search Projects").WebTable("Select Project to View").GetCellData (row,2)
	print row-1 & " of " & nRows-1 
	workPlanProfile = Browser("Search Projects").Page("Search Projects").WebTable("Select Project to View").ChildItemCount (row,3, "Image") @@ hightlight id_;_Browser("Search Projects").Page("Search Projects").Link("A/R Billing Upgrade")_;_script infofile_;_ZIP::ssf1.xml_;_
	If workPlanProfile = 1 Then

		Set projectLink = Browser("Search Projects").Page("Search Projects").WebTable("Select Project to View").ChildItem (row, 2, "Link", 0)
		projectLink.highlight
		projectLink.DoubleClick
			
		Browser("Search Projects_2").Page("Project Overview").WebElement("costTabNav").Click @@ hightlight id_;_Browser("Search Projects 2").Page("Project Overview").WebElement("costTabNav")_;_script infofile_;_ZIP::ssf63.xml_;_
		Browser("Search Projects_2").Page("Project Overview_2").WebElement("Work Plan Cost").Click @@ hightlight id_;_Browser("Search Projects 2").Page("Project Overview 2").WebElement("Work Plan Cost")_;_script infofile_;_ZIP::ssf64.xml_;_
		Browser("Search Projects_2").Page("Project Overview_2").Link("View details for the Actual").Click @@ script infofile_;_ZIP::ssf71.xml_;_
		TopLevelActualEffort = Browser("Search Projects_2").Page("View Work Plan_3").WebTable("In Planning").GetCellData (1,2)
		If TopLevelActualEffort = 0.0 Then 
			print "Project >" & projectName & "< does not have any effort"
 @@ hightlight id_;_Browser("Search Projects 2").Page("View Work Plan 5").WebElement("Green 2")_;_script infofile_;_ZIP::ssf88.xml_;_
		else
			print "Removing workaround work plan from project >" & projectName & "<"
			Browser("Search Projects_2").Page("View Work Plan_5").WebList("select").Select "Schedule View" @@ hightlight id_;_Browser("Search Projects 2").Page("View Work Plan 5").WebList("select")_;_script infofile_;_ZIP::ssf90.xml_;_
			wait 3 ' it takes a little while to change view
			
			
			if not Browser("View Work Plan_2").Page("View Work Plan").WebElement("workaround").Exist(4) then
				print "Project >" & projectName & "< does not have a workaround. Skipping..."
			else
			
				Browser("View Work Plan_2").Page("View Work Plan").WebElement("workaround").Click
				
				' This is an insight recording
				' needed to record the keystrokes
				Browser("View Work Plan_2").Page("View Work Plan").Link("Edit").Click

				'Browser("View Work Plan_2").InsightObject("InsightObject_3").Click @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf155.xml_;_
'				Browser("View Work Plan_2").InsightObject("InsightObject").Click
'				Window("Google Chrome").Type  micShiftDwn @@ hightlight id_;_394254_;_script infofile_;_ZIP::ssf156.xml_;_
'				Window("Google Chrome").Type micShiftDwn +  micTab  + micShiftUp @@ hightlight id_;_394254_;_script infofile_;_ZIP::ssf157.xml_;_
'				Window("Google Chrome").Type micShiftDwn +  micTab  + micShiftUp @@ hightlight id_;_394254_;_script infofile_;_ZIP::ssf158.xml_;_
'				Window("Google Chrome").Type  micShiftUp @@ hightlight id_;_394254_;_script infofile_;_ZIP::ssf159.xml_;_
'				Window("Google Chrome").Type  micDwn @@ hightlight id_;_394254_;_script infofile_;_ZIP::ssf160.xml_;_
'				Window("Google Chrome").Type  micDwn @@ hightlight id_;_394254_;_script infofile_;_ZIP::ssf161.xml_;_
'				Window("Google Chrome").Type  micReturn @@ hightlight id_;_394254_;_script infofile_;_ZIP::ssf162.xml_;_

				Window("Google Chrome").Type  micDwn
				Window("Google Chrome").Type  micDwn
				Window("Google Chrome").Type  micReturn
	
				Browser("Search Projects_2").Page("Enter Actuals").WebEdit("actualEffort").Set "0" @@ hightlight id_;_Browser("Search Projects 2").Page("Enter Actuals").WebEdit("actualEffort")_;_script infofile_;_ZIP::ssf85.xml_;_
				
				if Browser("Search Projects_2").Page("Enter Actuals").WebEdit("estimatedRemainingEffort").Exist (1) then ' some work plans do not have this visible @@ hightlight id_;_Browser("Search Projects 2").Page("Enter Actuals").WebEdit("estimatedRemainingEffort(61262")_;_script infofile_;_ZIP::ssf86.xml_;_
					Browser("Search Projects_2").Page("Enter Actuals").WebEdit("estimatedRemainingEffort").Set "0"
				end if			
				
				Browser("Search Projects_2").Page("Enter Actuals").WebEdit("percentComplete").Set "0" @@ hightlight id_;_Browser("Search Projects 2").Page("Enter Actuals").WebEdit("percentComplete")_;_script infofile_;_ZIP::ssf87.xml_;_
				Browser("Search Projects_2").Page("Enter Actuals").WebButton("Save").Click @@ hightlight id_;_Browser("Search Projects 2").Page("Enter Actuals").WebButton("Save")_;_script infofile_;_ZIP::ssf88.xml_;_
				Browser("Search Projects_2").Page("Enter Actuals").WebButton("Done").Click @@ hightlight id_;_Browser("Search Projects 2").Page("Enter Actuals").WebButton("Done")_;_script infofile_;_ZIP::ssf89.xml_;_
	
				' following 2 lines added Feb 12, 2020
				Browser("Search Projects_2").Page("Project Overview_2").WebElement("Work Plan Cost").Click @@ hightlight id_;_Browser("Search Projects 2").Page("Project Overview 2").WebElement("Work Plan Cost")_;_script infofile_;_ZIP::ssf64.xml_;_
				Browser("Search Projects_2").Page("Project Overview_2").Link("View details for the Actual").Click @@ script infofile_;_ZIP::ssf71.xml_;_
		
				Browser("Search Projects_2").Page("View Work Plan_5").WebElement("Workaround Task").Click @@ hightlight id_;_Browser("Search Projects 2").Page("View Work Plan 5").WebElement("schedule warning")_;_script infofile_;_ZIP::ssf96.xml_;_
				Browser("Search Projects_2").Page("View Work Plan_5").WebButton("deleteButton").Click @@ hightlight id_;_Browser("Search Projects 2").Page("View Work Plan 5").WebButton("deleteButton")_;_script infofile_;_ZIP::ssf97.xml_;_
				Browser("Search Projects_2").Page("View Work Plan_5").Frame("confirmationDelete").WebButton("positive.button").Click @@ hightlight id_;_Browser("Search Projects 2").Page("View Work Plan 5").Frame("confirmationDialogIF").WebButton("positive.button")_;_script infofile_;_ZIP::ssf98.xml_;_
				
				' verify that workaround is gone
				wait 5 ' it can take a little while. If we don't wait, the element might still exist
				' could do a sync on a window that says "Deleting..." but wait is easier
				if Browser("View Work Plan_2").Page("View Work Plan").WebElement("workaround").Exist (3) then
					msgbox "workaround not deleted from project " & projectName 			
				End If
				
				'set view back to Quick View
				print "Set view to Quick View  for project >" & projectName & "<"
				Browser("Search Projects_2").Page("View Work Plan_5").WebList("select").Select "Quick View"
				wait 3 ' it takes a little while to change view
			end if @@ hightlight id_;_Browser("Search Projects 2").Page("Enter Actuals").WebButton("Done")_;_script infofile_;_ZIP::ssf81.xml_;_
		End If

		Browser("Search Projects").Page("Project Overview_2").Link("Search Projects").Click @@ hightlight id_;_Browser("Search Projects").Page("Project Overview 2").Link("Search Projects")_;_script infofile_;_ZIP::ssf13.xml_;_
		
	else
		print "Project >" & projectName & "< does not have a work plan"
	End If
Next
 @@ hightlight id_;_Browser("Search Projects").Page("Project Overview 3").WebElement("pGanttNEW1574706060164")_;_script infofile_;_ZIP::ssf23.xml_;_
foo = 1
