print "Starting test: " & Environment("TestName")

'until no more browsers exist
While Browser("creationtime:=0").Exist(0)
'Close the browser
Browser("creationtime:=0").highlight
Browser("creationtime:=0").close
Wend

SystemUtil.Run "Chrome.exe", "http://ppmdemo:8084/itg/project/SearchProjects.do"

Browser("Search Projects_2").Page("Search Projects").WebEdit("USERNAME").Set "admin" @@ script infofile_;_ZIP::ssf86.xml_;_
Browser("Search Projects_2").Page("Search Projects").WebEdit("PASSWORD").SetSecure "61b375bbe33b17459618ababc7e1446da75aabb32eba3710" @@ script infofile_;_ZIP::ssf87.xml_;_
Browser("Search Projects_2").Page("Search Projects").WebElement("label-LOGON_SUBMIT_BUTTON_CAPT").Click @@ script infofile_;_ZIP::ssf88.xml_;_

' the following click always fails with a "can't find object message, but that is bull.
' the description is fine. I am convinced there is a corruption behind the scene. ron sercely
Browser("Search Projects_2").Page("Search Projects").Link("Search").Click @@ script infofile_;_ZIP::ssf85.xml_;_

nRows  = Browser("Search Projects").Page("Search Projects").WebTable("Select Project to View").RowCount
For row = 2 To nRows ' row 1 is the header row so skip
	projectName = Browser("Search Projects").Page("Search Projects").WebTable("Select Project to View").GetCellData (row,2)
	print row & " of " & nRows 
	workPlanProfile = Browser("Search Projects").Page("Search Projects").WebTable("Select Project to View").ChildItemCount (row,3, "Image") @@ hightlight id_;_Browser("Search Projects").Page("Search Projects").Link("A/R Billing Upgrade")_;_script infofile_;_ZIP::ssf1.xml_;_
	If workPlanProfile = 1 Then

		Set projectLink = Browser("Search Projects").Page("Search Projects").WebTable("Select Project to View").ChildItem (row, 2, "Link", 0)
		projectLink.highlight
		print "project " & projectlink.getroproperty("innertext")	& " has an href of " & projectlink.getroproperty("href")	
		projectLink.doubleClick

		Browser("Search Projects_2").Page("Project Overview").WebElement("costTabNav").Click @@ hightlight id_;_Browser("Search Projects 2").Page("Project Overview").WebElement("costTabNav")_;_script infofile_;_ZIP::ssf63.xml_;_
		Browser("Search Projects_2").Page("Project Overview_2").WebElement("Work Plan Cost").Click @@ hightlight id_;_Browser("Search Projects 2").Page("Project Overview 2").WebElement("Work Plan Cost")_;_script infofile_;_ZIP::ssf64.xml_;_
		Browser("Search Projects_2").Page("Project Overview_2").Link("View details for the Actual").Click
		Browser("Search Projects_2").Page("View Work Plan_3").WebTable("In Planning").Output CheckPoint("In Planning") @@ script infofile_;_ZIP::ssf71.xml_;_
		TopLevelActualEffort = Browser("Search Projects_2").Page("View Work Plan_3").WebTable("In Planning").GetCellData (1,2)
		If TopLevelActualEffort = 0.0 Then
			print "Project >" & projectName & "< does not have any effort"
		ElseIf "MSP Integration" = projectName Then ' skip because it has a MS Project integration, so does not allow new task
			print "Skipping >" & projectName & "< because of MS Project integration"
		
			else
				'		Browser("Search Projects_2").Page("Project Overview_2").Link("View details for the Actual").Click
				print "Adding workaround task to project >" & projectName & "<"
				Browser("Search Projects_2").Page("View Work Plan_3").WebElement("Seq 1").Click @@ hightlight id_;_Browser("Search Projects 2").Page("View Work Plan 3").WebElement("FunCollapse")_;_script infofile_;_ZIP::ssf68.xml_;_
				Browser("Search Projects_2").Page("View Work Plan_4").WebButton("insertButton").Click @@ hightlight id_;_Browser("Search Projects 2").Page("View Work Plan 4").WebButton("insertButton")_;_script infofile_;_ZIP::ssf72.xml_;_
				Browser("Search Projects_2").Page("Add Tasks Page").WebEdit("Task Name").Set "workaround"
				Browser("Search Projects_2").Page("Add Tasks Page").WebButton("Task Done").Click @@ hightlight id_;_Browser("Search Projects 2").Page("Add Tasks Page").WebButton("topDoneB")_;_script infofile_;_ZIP::ssf74.xml_;_

				If not Browser("Search Projects_2").Page("View Work Plan_5").Link("Edit").Exist (10) then ' we are probably in quick view, which does not allow editing
					Browser("Search Projects_2").Page("View Work Plan").WebList("select").Select "Schedule View" @@ script infofile_;_ZIP::ssf84.xml_;_
				end if

				Browser("Search Projects_2").Page("View Work Plan_5").Link("Edit").Click
				Browser("Search Projects_2").Page("View Work Plan_5").WebElement("edit.dropdown.actuals").Click
				Browser("Search Projects_2").Page("Enter Actuals").WebEdit("actualEffort").Set "10" @@ hightlight id_;_Browser("Search Projects 2").Page("Enter Actuals").WebEdit("actualEffort(6126260)")_;_script infofile_;_ZIP::ssf78.xml_;_
				Browser("Search Projects_2").Page("Enter Actuals").WebEdit("percentComplete").Set "10" @@ hightlight id_;_Browser("Search Projects 2").Page("Enter Actuals").WebEdit("percentComplete(6126260)")_;_script infofile_;_ZIP::ssf79.xml_;_
				Browser("Search Projects_2").Page("Enter Actuals").WebButton("Save").Click @@ hightlight id_;_Browser("Search Projects 2").Page("Enter Actuals").WebButton("Save")_;_script infofile_;_ZIP::ssf80.xml_;_
				Browser("Search Projects_2").Page("Enter Actuals").WebButton("Done").Click @@ hightlight id_;_Browser("Search Projects 2").Page("Enter Actuals").WebButton("Done")_;_script infofile_;_ZIP::ssf81.xml_;_
		End If

		For backToStart = 1 To 5 Step 1
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



foo = 1




