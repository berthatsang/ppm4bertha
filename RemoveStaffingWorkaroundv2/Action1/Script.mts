﻿print "Starting test: " & Environment("TestName")

'until no more browsers exist
While Browser("creationtime:=0").Exist(0)
	'Close the browser
	Browser("creationtime:=0").highlight
	Browser("creationtime:=0").close
Wend

SystemUtil.Run "Chrome.exe", "http://ppmdemo:8084/itg/project/SearchProjects.do"

Browser("Project Overview_2").Page("Search Projects").WebEdit("USERNAME").Set "admin" @@ script infofile_;_ZIP::ssf93.xml_;_
Browser("Project Overview_2").Page("Search Projects").WebEdit("PASSWORD").SetSecure "61b73ad87c686d0d7e0299e95eb63f9f2382f831ff5ed3df" @@ script infofile_;_ZIP::ssf94.xml_;_
Browser("Project Overview_2").Page("Search Projects").WebElement("label-LOGON_SUBMIT_BUTTON_CAPT").Click @@ script infofile_;_ZIP::ssf95.xml_;_

'Creating a description object

'Set btnsearch = Description.Create() 'Adding descriptions and properties
'btnsearch("type").value = "Button"
'btnsearch("name").value = "Search"
'btnsearch("html tag").value = "INPUT" ' Using the same to script it
'Browser("View Work Plan_2").Page("PPM Logon").WebButton(btnsearch).Click
Browser("Project Overview_2").Page("Search Projects").WebButton("SEARCH_BUTTON_LINK").Click

'Browser("Search Projects").Page("Search Projects").WebElement("html tag:=SPAN","innertext:=Search").click
'Browser("Project Overview_2").Page("Search Projects").WebElement("html tag:=SPAN","innertext:=Search", "x:=877", "y:=468").click
'Browser("Project Overview_2").Page("Search Projects").WebElement("html tag:=SPAN","innertext:=Search","innerhtml:=Search", "x:=877", "y:=468").click

'Browser("Project Overview_2").Page("Search Projects").WebElement("Search").Click
'Browser("Project Overview_2").InsightObject("SearchInsightObject").Click
'AIUtil("button", "", micFromBottom, 1).Click


' also, navigate back to "Dashboard - Key Status Information" using bread crumbs if you need to restart this script

Browser("Project Overview").Navigate "http://ppmdemo:8084/itg/project/SearchProjects.do"
Browser("Project Overview_2").Page("Search Projects").Link("SEARCH_2").Click @@ script infofile_;_ZIP::ssf79.xml_;_
Browser("Project Overview_2").Page("Search Projects").Link("Staffing Profiles").Click
'Browser("Project Overview_2").Page("Staffing Profile").Link("Search Staffing Profile").Click
Browser("Project Overview_2").Page("Search Staffing Profile").Image("Choose Status").Click @@ script infofile_;_ZIP::ssf41.xml_;_
 @@ script infofile_;_ZIP::ssf89.xml_;_
' add all status except baseline and 
Browser("Project Overview_2").Page("Search Staffing Profile").Frame("availFrame").WebElement("Active").FireEvent "onMouseOver" @@ hightlight id_;_Browser("Project Overview 2").Page("Search Staffing Profile").Frame("availFrame").WebElement("WebTable 2")_;_script infofile_;_ZIP::ssf56.xml_;_
Browser("Project Overview_2").Page("Search Staffing Profile").Frame("availFrame").WebElement("Active").click @@ hightlight id_;_Browser("Project Overview 2").Page("Search Staffing Profile").Frame("availFrame").WebElement("WebTable 2")_;_script infofile_;_ZIP::ssf56.xml_;_
Browser("Project Overview_2").Page("Search Staffing Profile").Frame("availFrame").WebElement("In Planning").FireEvent "onMouseOver" @@ hightlight id_;_Browser("Project Overview 2").Page("Search Staffing Profile").Frame("availFrame").WebElement("WebTable 3")_;_script infofile_;_ZIP::ssf57.xml_;_
Browser("Project Overview_2").Page("Search Staffing Profile").Frame("availFrame").WebElement("In Planning").Click @@ hightlight id_;_Browser("Project Overview 2").Page("Search Staffing Profile").Frame("availFrame").WebElement("WebTable 3")_;_script infofile_;_ZIP::ssf57.xml_;_
Browser("Project Overview_2").Page("Search Staffing Profile").Frame("availFrame").WebElement("Completed").FireEvent "onMouseOver" @@ hightlight id_;_Browser("Project Overview 2").Page("Search Staffing Profile").Frame("availFrame").WebElement("WebTable 4")_;_script infofile_;_ZIP::ssf58.xml_;_
Browser("Project Overview_2").Page("Search Staffing Profile").Frame("availFrame").WebElement("Completed").Click @@ hightlight id_;_Browser("Project Overview 2").Page("Search Staffing Profile").Frame("availFrame").WebElement("WebTable 4")_;_script infofile_;_ZIP::ssf58.xml_;_
Browser("Project Overview_2").Page("Search Staffing Profile").Frame("availFrame").WebElement("Cancelled").FireEvent "onMouseOver" @@ hightlight id_;_Browser("Project Overview 2").Page("Search Staffing Profile").Frame("availFrame").WebElement("WebTable 5")_;_script infofile_;_ZIP::ssf60.xml_;_
Browser("Project Overview_2").Page("Search Staffing Profile").Frame("availFrame").WebElement("Cancelled").Click @@ hightlight id_;_Browser("Project Overview 2").Page("Search Staffing Profile").Frame("availFrame").WebElement("WebTable 5")_;_script infofile_;_ZIP::ssf60.xml_;_
Browser("Project Overview_2").Page("Search Staffing Profile").Frame("availFrame").WebElement("Lock Down").FireEvent "onMouseOver" @@ hightlight id_;_Browser("Project Overview 2").Page("Search Staffing Profile").Frame("availFrame").WebElement("WebTable 6")_;_script infofile_;_ZIP::ssf61.xml_;_
Browser("Project Overview_2").Page("Search Staffing Profile").Frame("availFrame").WebElement("Lock Down").Click @@ hightlight id_;_Browser("Project Overview 2").Page("Search Staffing Profile").Frame("availFrame").WebElement("WebTable 6")_;_script infofile_;_ZIP::ssf61.xml_;_
Browser("Project Overview_2").Page("Search Staffing Profile").Frame("autoCompleteDialogIF").WebButton("btnOK").Click @@ script infofile_;_ZIP::ssf46.xml_;_
Browser("Project Overview_2").Page("Search Staffing Profile").WebEdit("pageSize").Set "500" @@ script infofile_;_ZIP::ssf47.xml_;_
Browser("Project Overview_2").Page("Search Staffing Profile").WebButton("PushButton_1581043286106_2773").Click @@ script infofile_;_ZIP::ssf48.xml_;_
 @@ script infofile_;_ZIP::ssf38.xml_;_
'Setting.webpackage("ReplayType") = 2 'Mouse
'Setting.webpackage("ReplayType") = 1 'Event

nRows  = Browser("Project Overview_2").Page("Search Staffing Profile").WebTable("Profile Name").RowCount
For row = 2 To nrows ' row 1 is the header row so skip
	print row -1 & " of " & nRows -1 
	set projectLink = Browser("Search Projects").Page("Search Projects").WebTable("Select Project to View").ChildItem (row, 1, "Link",0)
	projectLink.highlight
	projName = projectlink.getroproperty("innertext")
	projHref = projectlink.getroproperty("href")
	projectLink.DoubleClick
	If projName = "MSP Integration" Then
		print "skipping Staffing Profile " & projName & " because it is connected to MS Project"
	ElseIf projName = "Transaction Management" or projName = "Sparta System" Then
		print "skipping Staffing Profile " & projName & " because it has no staffing profile"
	ElseIf projName = "Barcode Asset Collection" or projName = "Sparta System" Then
		print "skipping Staffing Profile " & projName & " UNKOWN PROBLEM"
	ElseIf  Browser("Project Overview_2").Page("Staffing Profile").Link("workaround").Exist (3) then
		print "removing workaround from Staffing Profile " & projName & " which has an href of " & projHref
		wait 10
		Browser("Project Overview_2").Page("Staffing Profile").Link("workaround").highlight 
		wait 5 ' otherwise, the following click sometimes get "lost"
		' I think this is because it is using an Insight object
		'Browser("Project Overview_2").InsightObject("workaroundCheckpox1").Click
		Browser("Project Overview_2").InsightObject("workaround").Click

'Browser("Project Overview_2").Page("Staffing Profile").WebElement("WebElement_5").Click
'Browser("Project Overview_2").Page("Staffing Profile").WebElement("WebElement_4").Click @@ script infofile_;_ZIP::ssf96.xml_;_

		Browser("Project Overview_2").Page("Staffing Profile").WebElement("confirm Delete").Click @@ script infofile_;_ZIP::ssf94.xml_;_
		Browser("Project Overview_2").Page("Staffing Profile").Link("Yes").Click @@ script infofile_;_ZIP::ssf95.xml_;_
		
	else
		print "project " & projName	& " does not have a workaround, but should"
end if
	Browser("Project Overview_2").Page("Staffing Profile").Link("Search Staffing Profile").Click
Next



foo = 1

