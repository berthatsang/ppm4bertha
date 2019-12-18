nRows  = Browser("Search Projects").Page("Search Projects").WebTable("Select Project to View").RowCount
For row = 2 To nRows ' row 1 is the header row so skip
	projectName = Browser("Search Projects").Page("Search Projects").WebTable("Select Project to View").GetCellData (row,2)
	print row & " of " & nRows 
	workPlanProfile = Browser("Search Projects").Page("Search Projects").WebTable("Select Project to View").ChildItemCount (row,3, "Image") @@ hightlight id_;_Browser("Search Projects").Page("Search Projects").Link("A/R Billing Upgrade")_;_script infofile_;_ZIP::ssf1.xml_;_
	If workPlanProfile = 1 Then

		Set projectLink = Browser("Search Projects").Page("Search Projects").WebTable("Select Project to View").ChildItem (row, 2, "Link", 0)
		projectLink.highlight
		projectLink.Click
			
		Browser("Search Projects_2").Page("Project Overview").WebElement("costTabNav").Click @@ hightlight id_;_Browser("Search Projects 2").Page("Project Overview").WebElement("costTabNav")_;_script infofile_;_ZIP::ssf63.xml_;_
		Browser("Search Projects_2").Page("Project Overview_2").WebElement("Work Plan Cost").Click @@ hightlight id_;_Browser("Search Projects 2").Page("Project Overview 2").WebElement("Work Plan Cost")_;_script infofile_;_ZIP::ssf64.xml_;_
		Browser("Search Projects_2").Page("Project Overview_2").Link("View details for the Actual").Click
		Browser("Search Projects_2").Page("View Work Plan_3").WebTable("In Planning").Output CheckPoint("In Planning") @@ script infofile_;_ZIP::ssf71.xml_;_
		TopLevelActualEffort = Browser("Search Projects_2").Page("View Work Plan_3").WebTable("In Planning").GetCellData (1,2)
		If TopLevelActualEffort = 0.0 Then 
			print "Project >" & projectName & "< does not have any effort"
		else
'		Browser("Search Projects_2").Page("Project Overview_2").Link("View details for the Actual").Click @@ hightlight id_;_Browser("Search Projects 2").Page("View Work Plan 3").WebElement("FunCollapse")_;_script infofile_;_ZIP::ssf67.xml_;_
			print "Adding workaround task to project >" & projectName & "<"
			Browser("Search Projects_2").Page("View Work Plan_3").WebElement("Seq 1").Click @@ hightlight id_;_Browser("Search Projects 2").Page("View Work Plan 3").WebElement("FunCollapse")_;_script infofile_;_ZIP::ssf68.xml_;_
			Browser("Search Projects_2").Page("View Work Plan_4").WebButton("insertButton").Click @@ hightlight id_;_Browser("Search Projects 2").Page("View Work Plan 4").WebButton("insertButton")_;_script infofile_;_ZIP::ssf72.xml_;_
			Browser("Search Projects_2").Page("Add Tasks Page").WebEdit("Task Name").Set "workaround"
			Browser("Search Projects_2").Page("Add Tasks Page").WebButton("Task Done").Click @@ hightlight id_;_Browser("Search Projects 2").Page("Add Tasks Page").WebButton("topDoneB")_;_script infofile_;_ZIP::ssf74.xml_;_
'			Browser("Search Projects_2").Page("View Work Plan_5").Image("dbtn-arrow-active").Click
			Browser("Search Projects_2").Page("View Work Plan_5").Link("Edit").Click @@ hightlight id_;_Browser("Search Projects 2").Page("View Work Plan 5").Link("Edit")_;_script infofile_;_ZIP::ssf77.xml_;_
			Browser("Search Projects_2").Page("View Work Plan_5").WebElement("edit.dropdown.actuals").Click
			Browser("Search Projects_2").Page("Enter Actuals").WebEdit("actualEffort").Set "10" @@ hightlight id_;_Browser("Search Projects 2").Page("Enter Actuals").WebEdit("actualEffort(6126260)")_;_script infofile_;_ZIP::ssf78.xml_;_
			Browser("Search Projects_2").Page("Enter Actuals").WebEdit("percentComplete").Set "10" @@ hightlight id_;_Browser("Search Projects 2").Page("Enter Actuals").WebEdit("percentComplete(6126260)")_;_script infofile_;_ZIP::ssf79.xml_;_
			Browser("Search Projects_2").Page("Enter Actuals").WebButton("Save").Click @@ hightlight id_;_Browser("Search Projects 2").Page("Enter Actuals").WebButton("Save")_;_script infofile_;_ZIP::ssf80.xml_;_
			Browser("Search Projects_2").Page("Enter Actuals").WebButton("Done").Click @@ hightlight id_;_Browser("Search Projects 2").Page("Enter Actuals").WebButton("Done")_;_script infofile_;_ZIP::ssf81.xml_;_
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




