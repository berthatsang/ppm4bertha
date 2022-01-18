print "Starting test:" & environment("TestName")

'close browsers until no more browsers exist
While Browser("creationtime:=0").Exist(0)
'Close the browser
Browser("creationtime:=0").highlight
Browser("creationtime:=0").close
wait 3
Wend

' recommend you use chrome - seems to replay faster
SystemUtil.Run "chrome.exe", "http://ppmdemo:8084/itg"

Browser("Scenario Details").Page("PPM Logon").WebElement("Username").Click @@ hightlight id_;_Browser("Scenario Details").Page("PPM Logon").WebElement("Username")_;_script infofile_;_ZIP::ssf2.xml_;_
Browser("Scenario Details").Page("PPM Logon").WebEdit("USERNAME").Set "admin" @@ hightlight id_;_Browser("Scenario Details").Page("PPM Logon").WebEdit("USERNAME")_;_script infofile_;_ZIP::ssf3.xml_;_
Browser("Scenario Details").Page("PPM Logon").WebEdit("PASSWORD").Set "mfDemo$20" @@ hightlight id_;_Browser("Scenario Details").Page("PPM Logon").WebEdit("PASSWORD")_;_script infofile_;_ZIP::ssf4.xml_;_
'Browser("Scenario Details").Page("PPM Logon").WebEdit("PASSWORD").SetSecure "61b375bbe33b17459618ababc7e1446da75aabb32eba3710"
Browser("Scenario Details").Page("PPM Logon").WebElement("Sign in Button").Click @@ hightlight id_;_Browser("Scenario Details").Page("PPM Logon").WebElement("label-LOGON SUBMIT BUTTON CAPT")_;_script infofile_;_ZIP::ssf5.xml_;_

Window("Google Chrome").Type micCtrlDwn + "0" + micCtrlUp ' set zoom to 100% @@ hightlight id_;_1509550_;_script infofile_;_ZIP::ssf57.xml_;_
Const ZOOM_LEVEL = 8
For zoom = 1 To ZOOM_LEVEL Step 1
	Window("Google Chrome").Type micCtrlDwn + "-" + micCtrlUp
	wait 1
Next

Browser("Scenario Details").Page("PPM Logon").Link("OPEN").Click @@ hightlight id_;_Browser("Scenario Details").Page("PPM Logon").Link("OPEN")_;_script infofile_;_ZIP::ssf6.xml_;_
Browser("Scenario Details").Page("PPM Logon").Link("What-if Analysis").Click @@ hightlight id_;_Browser("Scenario Details").Page("PPM Logon").Link("What-if Analysis")_;_script infofile_;_ZIP::ssf7.xml_;_

'Screen scraping to copy the values that we'll use later
' now start processing the plan
' the scenario name is in the data table @@ hightlight id_;_Browser("Scenario Details").Page("Scenario List").Link("Current Plan")_;_script infofile_;_ZIP::ssf8.xml_;_
' it is parameterized in the link description in the OR
if Browser("Scenario Details").Page("Scenario List").Link("Scenario to Work On").Exist (30) then 
	Browser("Scenario Details").Page("Scenario List").Link("Scenario to Work On").Click 
end if


if Browser("Scenario Details").Page("Edit Details").WebButton("Apply Preview").Exist (90) then ' takes a long tiime @@ hightlight id_;_Browser("Scenario Details").Page("Scenario Details").WebButton("Apply Preview")_;_script infofile_;_ZIP::ssf9.xml_;_
end if

' in-icon is blue
'out-icon  in is grey
Set blue = description.Create
blue("Title").value = "Move out"
Set blueLabel = description.Create
blueLabel("href").value = ".*/itg/web/knta/crt/RequestDetail\.jsp\?REQUEST_ID=\d*"

'113 is arbitrarily large enough
For curBlueRow = 0 To 113 Step 1 
	blue("Index").value = curBlueRow
	'this is how we know we're done	
	If (	not Browser("Scenario Details").Page("Edit Details").WebButton(blue).Exist (5) ) Then
		Exit for
	End If

	' Uncomment this if you want debug statements
	' foo = 1
	
	if foo = 1 then 'debug
		print "curBlueRow: " & curBlueRow
		Browser("Scenario Details").Page("Edit Details").WebButton(blue).highlight
	end if

  	blueLabel("Index").value = curBlueRow
	if foo = 1 then
		Browser("Scenario Details").Page("Edit Details").Link(blueLabel).Highlight
	end if 
	
	curContents =  Browser("Scenario Details").Page("Edit Details").Link(blueLabel).GetROProperty ("text") 
	blueContents = blueContents & curContents &","
	'	Browser("Scenario Details").Page("Edit Details").WebButton(blue).Highlight

	if foo = 1 then
		print "Included Contents:" & curContents
	end if
	
next

print "All Included Contents:" & blueContents

Set grey = description.Create
grey("Title").value = "Move in"
Set greyLabel = description.Create
greyLabel("href").value = ".*/itg/web/knta/crt/RequestDetail\.jsp\?REQUEST_ID=\d*"

For curRow = 0 To 200 Step 1 
	curGreyRow = curRow
	grey("Index").value = curRow
	If (	not Browser("Scenario Details").Page("Edit Details").WebButton(grey).Exist (5) ) Then
		Exit for
	End If

	greyLabel("Index").value = curGreyRow + curBlueRow
	curContents = Browser("Scenario Details").Page("Edit Details").Link(greyLabel).GetROProperty ("text") &","
	greyContents = greyContents & curContents
'	Browser("Scenario Details").Page("Edit Details").WebButton(grey).Highlight
	
	If foo = 1 Then
		print "Excluded Content" & curContents
		print "All excluded content" & greyContents
	
	End If
next
print "All excluded content" & greyContents

' get the details of the currently selected scenario - so we can use them in the new scenario @@ hightlight id_;_Browser("Scenario Details").Page("Scenario List").Link("Current Plan")_;_script infofile_;_ZIP::ssf29.xml_;_
Browser("Scenario Details").Page("Edit Details").WebButton("Pencil Icon").Click ' Note - it is the pencil icon
planName = Browser("Scenario Details").Page("Change Scenario Constraints").WebEdit("scenarioName").GetROProperty ("value")
planStartDate = Browser("Scenario Details").Page("Change Scenario Constraints").WebEdit("Scenario Start Date").GetROProperty ("value")
planEndDate = Browser("Scenario Details").Page("Change Scenario Constraints").WebEdit("Scenario End Date").GetROProperty ("value")

wait 3
Window("Google Chrome").Type micCtrlDwn + "0" + micCtrlUp ' set zoom to 100% @@ hightlight id_;_1509550_;_script infofile_;_ZIP::ssf57.xml_;_

'Is it a Portfolio or a Program?
If  Browser("Scenario Details").InsightObject("enabledSelectPortfolio").Exist (10) then
	PortfolioOrProgam = "Portfolio"
	planPortfolio = Browser("Scenario Details").Page("Change Scenario Constraints").WebEdit("SelectPortfolio").GetROProperty("value")
else
	PortfolioOrProgam = "Program"
	planPortfolio = Browser("Scenario Details").Page("Change Scenario Constraints").WebEdit("SelectProgram").GetROProperty("value")
End If
For zoom = 1 To ZOOM_LEVEL Step 1
	Window("Google Chrome").Type micCtrlDwn + "-" + micCtrlUp
	wait 1
Next
        
planBudget = Browser("Scenario Details").Page("Change Scenario Constraints").WebEdit("totalBudget").GetROProperty ("value")

' Look at Resources Supply and Resource Pools set
Set resourceDesc = Description.Create()
resourceDesc("class").value = "acl-select-item-desc"
resourceDesc("html tag").value = "DIV"
For xxx = 0 To 10 Step 1
	resourceDesc("index").value = xxx
	if Browser("Scenario Details").Page("Create Scenario").WebElement(resourceDesc).Exist (3) then
		Browser("Scenario Details").Page("Create Scenario").WebElement(resourceDesc).Highlight
		title = Browser("Scenario Details").Page("Create Scenario").WebElement(resourceDesc).GetROProperty("Title")
		print "Resource Pool: " & title
		resources = resources & title &","
	else
		Exit for
	end if 
next
print "All Resource Pools:" & resources

Browser("Scenario Details").Page("Change Scenario Constraints").WebButton("Cancel").Click @@ hightlight id_;_Browser("Scenario Details").Page("Change Scenario Constraints").WebButton("Cancel")_;_script infofile_;_ZIP::ssf18.xml_;_
Browser("Scenario Details").Page("Edit Details").WebButton("Back To Scenario List").Click @@ hightlight id_;_Browser("Scenario Details").Page("Edit Details").WebButton("Back To Scenario List")_;_script infofile_;_ZIP::ssf30.xml_;_

' ok, got the old values - now create new scenarios
Browser("Scenario Details").Page("Scenario List").WebButton("Create").Click ' new scenario @@ hightlight id_;_Browser("Scenario Details").Page("Scenario List").WebButton("Create")_;_script infofile_;_ZIP::ssf19.xml_;_
if Browser("Scenario Details").Page("Create Scenario").WebEdit("scenarioName").Exist (90) then
end if 
Browser("Scenario Details").Page("Create Scenario").WebEdit("scenarioName").Set datatable.Value("Scenarios") & " RRS" @@ hightlight id_;_Browser("Scenario Details").Page("Create Scenario").WebEdit("scenarioName")_;_script infofile_;_ZIP::ssf20.xml_;_

' update Dates to next year
fieldVal = split (planStartDate, " ")
fieldVal(1) = cInt(fieldVal(1))+1
planStartDate = fieldVal(0) &" "& fieldVal(1)
fieldVal = split (planEndDate, " ")
fieldVal(1) = cInt(fieldVal(1))+1
planEndDate = fieldVal(0) &" "& fieldVal(1)

' Test data for debugging purposes
'planStartDate = "Dec 22"
'planEndDate = "Dec 24"
' Waits needed to ensure successful input of dates
Browser("Scenario Details").Page("Change Scenario Constraints").WebEdit("Scenario Start Date").Click
wait 2
Window("Google Chrome").Type planStartDate @@ hightlight id_;_1640732_;_script infofile_;_ZIP::ssf76.xml_;_
wait 2
Window("Google Chrome").Type micTab

Browser("Scenario Details").Page("Change Scenario Constraints").WebEdit("Scenario End Date").Click
wait 2
Window("Google Chrome").Type planEndDate
wait 2
Window("Google Chrome").Type  micTab

' Test data for debugging purposes
'planPortfolio = "ALM Octane"
'planPortfolio = "Black Diamond Initiative"
'PortfolioOrProgam = "Program"
if PortfolioOrProgam = "Portfolio" then
	Browser("Scenario Details").Page("Change Scenario Constraints").WebEdit("SelectPortfolio_New").click
	wait 2
	Window("Google Chrome").Type  planPortfolio 
	wait 2
	Window("Google Chrome").Type  micTab @@ hightlight id_;_1442356_;_script infofile_;_ZIP::ssf65.xml_;_
else
	Browser("Scenario Details").Page("Create Scenario").WebElement("program").Click
	Browser("Scenario Details").Page("Change Scenario Constraints").WebEdit("SelectProgram").click
	wait 2
	Window("Google Chrome").Type  planPortfolio 
	wait 2
	Window("Google Chrome").Type  micTab @@ hightlight id_;_1442356_;_script infofile_;_ZIP::ssf65.xml_;_
End If

Browser("Scenario Details").Page("Change Scenario Constraints").WebEdit("totalBudget").Set planBudget

' Add Resources Supply - need to scroll down on the pop-up list of Resource Pools to find correct one
Browser("Scenario Details").Page("Create Scenario").WebButton("+ Add Resource").Click
' make sure that the +Add Resource is visible, if not click it again
if not Browser("Scenario Details").InsightObject("Zoomed Next Resource Page").Exist then @@ hightlight id_;_5_;_script infofile_;_ZIP::ssf50.xml_;_
	wait 5
	Browser("Scenario Details").Page("Create Scenario").WebButton("+ Add Resource").Click
end if

' Test data for debugging purposes
'resources = "Non-Developers (WW) (FS),Offshore Partner A (FS),UI and Web Development - Team 1 (AMS) (FS),"
resourceArray = split (resources, ",")
For each resource in resourceArray
	If resource = "" Then
		Exit for
	End If

	'Set property value in the OR
	Browser("Scenario Details").Page("Create Scenario").WebElement("Resource Select Element").SetTOProperty "innertext",resource @@ hightlight id_;_Browser("Scenario Details").Page("Create Scenario").WebElement("AOS Team")_;_script infofile_;_ZIP::ssf46.xml_;_
	For tries = 1 To 10 Step 1 ' ugly - got to scroll until we find it
		if Browser("Scenario Details").Page("Create Scenario").WebElement("Resource Select Element").Exist(1) then @@ hightlight id_;_Browser("Scenario Details").Page("Create Scenario").WebElement("AOS Team")_;_script infofile_;_ZIP::ssf46.xml_;_
			Exit for
		end if
			Browser("Scenario Details").InsightObject("Zoomed Next Resource Page").Highlight 'scrollbar
			Browser("Scenario Details").InsightObject("Zoomed Next Resource Page").Click ' tab down @@ hightlight id_;_5_;_script infofile_;_ZIP::ssf50.xml_;_
	Next
	Browser("Scenario Details").Page("Create Scenario").WebElement("Resource Select Element").Click
	Window("Google Chrome").Type  micTab @@ hightlight id_;_1442356_;_script infofile_;_ZIP::ssf65.xml_;_
	wait 2
	Browser("Scenario Details").Page("Create Scenario").WebButton("+ Add Resource").Click
next


Browser("Scenario Details").Page("Create Scenario").WebButton("Create").Click
if Browser("Scenario Details").Page("Scenario Details").WebButton("Apply Preview").Exist (30) then
else
	msgbox "Scenario not created:" & datatable("Scenarios").value
end if
wait 5

itemsToMakeGrey = split (greyContents , ",")
' if running from here - you need to populate the value of greyContents
'itemsToMakeGrey = split  ("2-Hour Delivery,,ACME Project StreamPath,,Advantage Airline Merger,,Black Diamond Initiative,,Compliance Tracker,,Corporate Help Desk,,Corporate Intranet,,Covid-19 Website,,CRM One World,,Customer Dynamics,,Data Visualization,,Delphyne Craft Epic,,Enterprise Applications,,Enterprise Information Portal,,Idea Sharing,,Legacy Data Warehouse,,Location intelligence,,LTO System,,Map Mash-Up,,Master KPI System,,Mobile Office,,Neptune II,,Network Security Update,,NextGen NPI,,PPM Phase II,,Predictive Analytics,,Product Configurator,,Single Sign On,,Siren Craft Epic,,Smartphone Calendar,,SOX Audit Tracker,,Strategic Sourcing,,Text Analytics,,Vendor Management Web Site,,VNS System Update", ",")

Set moveOut = description.Create

moveOut("Title").value = "Move out"
Set moveLabel = description.Create
moveLabel("href").value = ".*/itg/web/knta/crt/RequestDetail\.jsp\?REQUEST_ID=\d*"

environment ("continueRow") = 0
 
For each item in itemsToMakeGrey
	If item = "" Then ' there is a blank at the end because of the way the concatenation was done
		Exit for
	End If 		
	startingRow = environment ("continueRow")
	For curGreyRow = startingRow To 55 Step 1 
		moveLabel("Index").value = curGreyRow
		curContents = Browser("Scenario Details").Page("Edit Details").Link(moveLabel).GetROProperty ("text")
		If item = curContents Then 'we found it. Click on the check
			moveOut("Index").value = curGreyRow
			Browser("Scenario Details").Page("Edit Details").WebButton(moveOut).Click
			environment ("continueRow") = curGreyRow
			Exit for ' time to look for next label
		End If
	Next
	
Next

Browser("Scenario Details").Page("Scenario Details").WebButton("Save").Click @@ script infofile_;_ZIP::ssf83.xml_;_
if Browser("Scenario Details").Page("Scenario Details").WebButton("Apply Preview").Exist (30) then
else
	msgbox "Scenario not created"
end if @@ hightlight id_;_Browser("Scenario Details").Page("Edit Details").WebButton("Back To Scenario List")_;_script infofile_;_ZIP::ssf26.xml_;_
Browser("Scenario Details").Page("Edit Details").WebButton("Back To Scenario List").Click @@ hightlight id_;_Browser("Scenario Details").Page("Edit Details").WebButton("Back To Scenario List")_;_script infofile_;_ZIP::ssf26.xml_;_

Window("Google Chrome").Type micCtrlDwn + "0" + micCtrlUp ' set zoom to 100% @@ hightlight id_;_1509550_;_script infofile_;_ZIP::ssf57.xml_;_

foo = 1
