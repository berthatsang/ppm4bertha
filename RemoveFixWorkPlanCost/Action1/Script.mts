nRows  = Browser("Search Projects").Page("Search Projects").WebTable("Select Project to View").RowCount
For row = 10 To nRows ' row 1 is the header row so skip
	projectName = Browser("Search Projects").Page("Search Projects").WebTable("Select Project to View").GetCellData (row,2)
'	print row & " of " & nRows 
	workPlanProfile = Browser("Search Projects").Page("Search Projects").WebTable("Select Project to View").ChildItemCount (row,3, "Image") @@ hightlight id_;_Browser("Search Projects").Page("Search Projects").Link("A/R Billing Upgrade")_;_script infofile_;_ZIP::ssf1.xml_;_
	If workPlanProfile = 1 Then

		Set projectLink = Browser("Search Projects").Page("Search Projects").WebTable("Select Project to View").ChildItem (row, 2, "Link", 0)
		projectLink.highlight
		projectLink.Click
			
		Browser("Search Projects_2").Page("Project Overview").WebElement("costTabNav").Click @@ hightlight id_;_Browser("Search Projects 2").Page("Project Overview").WebElement("costTabNav")_;_script infofile_;_ZIP::ssf63.xml_;_
		Browser("Search Projects_2").Page("Project Overview_2").WebElement("Work Plan Cost").Click @@ hightlight id_;_Browser("Search Projects 2").Page("Project Overview 2").WebElement("Work Plan Cost")_;_script infofile_;_ZIP::ssf64.xml_;_
		Browser("Search Projects_2").Page("Project Overview_2").Link("View details for the Actual").Click @@ script infofile_;_ZIP::ssf71.xml_;_
		TopLevelActualEffort = Browser("Search Projects_2").Page("View Work Plan_3").WebTable("In Planning").GetCellData (1,2)
		If TopLevelActualEffort = 0.0 Then 
			print "Project >" & projectName & "< does not have any effort"
 @@ hightlight id_;_Browser("Search Projects 2").Page("View Work Plan 5").WebElement("Green 2")_;_script infofile_;_ZIP::ssf88.xml_;_
		else
			print "Removing workaround task to project >" & projectName & "<"
			Browser("Search Projects_2").Page("View Work Plan_5").WebList("select").Select "Schedule View" @@ hightlight id_;_Browser("Search Projects 2").Page("View Work Plan 5").WebList("select")_;_script infofile_;_ZIP::ssf90.xml_;_
			Browser("Search Projects_2").Page("View Work Plan_5").WebTable("WorkPlan").Check CheckPoint("workaround task exists") @@ hightlight id_;_Browser("Search Projects 2").Page("View Work Plan 5").WebTable("0")_;_script infofile_;_ZIP::ssf133.xml_;_
			Browser("Search Projects_2").Page("View Work Plan_5").WebElement("Workaround Task").Click	 
			Browser("Search Projects_2").Page("View Work Plan_5").Link("Edit").Click @@ hightlight id_;_Browser("Search Projects 2").Page("View Work Plan 5").Link("Edit")_;_script infofile_;_ZIP::ssf84.xml_;_
			
Browser("Search Projects_2").Page("View Work Plan_6").Link("Edit").Click @@ hightlight id_;_Browser("Search Projects 2").Page("View Work Plan 6").Link("Edit")_;_script infofile_;_ZIP::ssf99.xml_;_
Browser("Search Projects_2").Page("View Work Plan_7").Link("Edit").Click @@ hightlight id_;_Browser("Search Projects 2").Page("View Work Plan 7").Link("Edit")_;_script infofile_;_ZIP::ssf100.xml_;_
			Browser("Search Projects_2").Page("Enter Actuals").WebEdit("actualEffort").Set "0" @@ hightlight id_;_Browser("Search Projects 2").Page("Enter Actuals").WebEdit("actualEffort")_;_script infofile_;_ZIP::ssf85.xml_;_
			Browser("Search Projects_2").Page("Enter Actuals").WebEdit("estimatedRemainingEffort").Set "0" @@ hightlight id_;_Browser("Search Projects 2").Page("Enter Actuals").WebEdit("estimatedRemainingEffort(61262")_;_script infofile_;_ZIP::ssf86.xml_;_
			Browser("Search Projects_2").Page("Enter Actuals").WebEdit("percentComplete").Set "0" @@ hightlight id_;_Browser("Search Projects 2").Page("Enter Actuals").WebEdit("percentComplete")_;_script infofile_;_ZIP::ssf87.xml_;_
			Browser("Search Projects_2").Page("Enter Actuals").WebButton("Save").Click @@ hightlight id_;_Browser("Search Projects 2").Page("Enter Actuals").WebButton("Save")_;_script infofile_;_ZIP::ssf88.xml_;_
			Browser("Search Projects_2").Page("Enter Actuals").WebButton("Done").Click @@ hightlight id_;_Browser("Search Projects 2").Page("Enter Actuals").WebButton("Done")_;_script infofile_;_ZIP::ssf89.xml_;_
			Browser("Search Projects_2").Page("Enter Actuals").WebButton("Save").Click @@ hightlight id_;_Browser("Search Projects_2").Page("Enter Actuals")_;_script infofile_;_ZIP::ssf92.xml_;_
'			Browser("Search Projects_2").Back @@ hightlight id_;_6095714_;_script infofile_;_ZIP::ssf93.xml_;_
			Browser("Search Projects_2").Page("View Work Plan_5").WebElement("Workaround Task").Click @@ hightlight id_;_Browser("Search Projects 2").Page("View Work Plan 5").WebElement("schedule warning")_;_script infofile_;_ZIP::ssf96.xml_;_
			Browser("Search Projects_2").Page("View Work Plan_5").WebButton("deleteButton").Click @@ hightlight id_;_Browser("Search Projects 2").Page("View Work Plan 5").WebButton("deleteButton")_;_script infofile_;_ZIP::ssf97.xml_;_
			Browser("Search Projects_2").Page("View Work Plan_5").Frame("confirmationDelete").WebButton("positive.button").Click @@ hightlight id_;_Browser("Search Projects 2").Page("View Work Plan 5").Frame("confirmationDialogIF").WebButton("positive.button")_;_script infofile_;_ZIP::ssf98.xml_;_
			
			' verify that workaround is gone
			Row2TaskName = Browser("Search Projects_2").Page("View Work Plan_5").WebTable("WorkPlan").GetCellData (2,5)
			If Row2TaskName = "workaround" Then
				msgbox "workaround not deleted from project " & projectName 			
			End If
			 @@ hightlight id_;_Browser("Search Projects 2").Page("Enter Actuals").WebButton("Done")_;_script infofile_;_ZIP::ssf81.xml_;_
		End If

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
	else
		print "Project >" & projectName & "< does not have a work plan"
	End If
Next


 @@ hightlight id_;_Browser("Search Projects").Page("Project Overview 3").WebElement("pGanttNEW1574706060164")_;_script infofile_;_ZIP::ssf23.xml_;_
foo = 1




