' Note: Object recognition seems to work much better if using IE as the browser

nRows  = Browser("Search Projects").Page("Search Projects").WebTable("Select Project to View").RowCount
For row = 2 To nRows ' row 1 is the header row so skip
	print row & " of " &  nRows
	staffingProfile = Browser("Search Projects").Page("Search Projects").WebTable("Select Project to View").ChildItemCount (row, 4, "Image") @@ hightlight id_;_Browser("Search Projects").Page("Search Projects").Link("A/R Billing Upgrade")_;_script infofile_;_ZIP::ssf1.xml_;_
	If staffingProfile = 1 Then
        Set projectLink = Browser("Search Projects").Page("Search Projects").WebTable("Select Project to View").ChildItem (row, 2, "Link", 0)
		projectLink.highlight
		projectLink.Click
		Browser("Project Overview").Page("Project Overview").WebElement("staffingTabNav").Click @@ hightlight id_;_Browser("Search Projects").Page("Project Overview").WebElement("staffingTabNav")_;_script infofile_;_ZIP::ssf2.xml_;_

		if Browser("Project Overview").InsightObject("woraround checkbox").Exist (30) then
			wait 2
			Browser("Project Overview").InsightObject("woraround checkbox").Click
		end if
		
		if Browser("Project Overview").Page("Project Overview_3").Link("Confirm").Exist (3) then @@ hightlight id_;_Browser("Project Overview").Page("Project Overview 3").Link("Confirm")_;_script infofile_;_ZIP::ssf31.xml_;_
			Browser("Project Overview").Page("Project Overview_3").Link("Confirm").Click @@ hightlight id_;_Browser("Project Overview").Page("Project Overview 3").Link("Confirm")_;_script infofile_;_ZIP::ssf31.xml_;_
			Browser("Project Overview").InsightObject("woraround checkbox").Click
		end if
		
		Browser("Project Overview").Page("Project Overview_2").WebElement("delete staff").Click
		Browser("Project Overview").Page("Project Overview_2").Link("Yes").Click @@ hightlight id_;_Browser("Project Overview").Page("Project Overview 2").Link("Yes")_;_script infofile_;_ZIP::ssf30.xml_;_
		
		For backToStart = 1 To 5 Step 1
			
			if Browser("Project Overview").Page("Project Overview").Link("Search Projects").Exist (10) then ' standard hack of having to wait after the object appears
				Browser("Project Overview").Page("Project Overview").Link("Search Projects").Click @@ hightlight id_;_Browser("Search Projects").Page("Project Overview 2").Link("Search Projects")_;_script infofile_;_ZIP::ssf13.xml_;_
				wait 5
				If Browser("Search Projects").Page("Search Projects").WebTable("Select Project to View").exist (2) Then
					Exit for
				end if
			End If
		next
	end if
Next


 @@ hightlight id_;_Browser("Search Projects").Page("Project Overview 3").WebElement("pGanttNEW1574706060164")_;_script infofile_;_ZIP::ssf23.xml_;_
foo = 1

