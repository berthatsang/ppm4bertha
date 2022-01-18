print "Checking for: " & DataTable("p_scenarioName", dtLocalSheet)
If Browser("Scenario List").Page("Scenario List").Link("Scenario Name RRS").Exist(3) Then
	
	print "Deleting: " & DataTable("p_scenarioName", dtLocalSheet)
	
	'Delete the old scenario
	Browser("Scenario List").Page("Scenario List").Link("Scenario Name").Click
	Browser("Scenario List").Page("Scenario Details").WebButton("Delete").Click @@ hightlight id_;_Browser("Scenario List").Page("Scenario Details 2").WebButton("Delete")_;_script infofile_;_ZIP::ssf10.xml_;_
	Browser("Scenario List").Page("Scenario Details").WebButton("Yes").Click @@ hightlight id_;_Browser("Scenario List").Page("Scenario Details 2").WebButton("Yes")_;_script infofile_;_ZIP::ssf11.xml_;_
	
	'Rename the new scenario
	Browser("Scenario List").Page("Scenario List").Link("Scenario Name RRS").Click @@ hightlight id_;_Browser("Scenario List").Page("Scenario List").Link("Black Diamond Scenario")_;_script infofile_;_ZIP::ssf4.xml_;_
	Browser("Scenario List").Page("Scenario Details").WebButton("EditButton").Click @@ hightlight id_;_Browser("Scenario List").Page("Scenario Details").WebButton("WebButton")_;_script infofile_;_ZIP::ssf5.xml_;_
	
	Browser("Scenario List").Page("Change Scenario Constraints").WebEdit("scenarioName").Set DataTable("p_scenarioName", dtLocalSheet)
	
	Browser("Scenario List").Page("Change Scenario Constraints").WebButton("Try It Out Button").Click @@ hightlight id_;_Browser("Scenario List").Page("Change Scenario Constraints").WebButton("Try It Out")_;_script infofile_;_ZIP::ssf7.xml_;_
	Browser("Scenario List").Page("Scenario Details").WebButton("Save Button").Click @@ hightlight id_;_Browser("Scenario List").Page("Scenario Details").WebButton("Save")_;_script infofile_;_ZIP::ssf8.xml_;_
	Browser("Scenario List").Page("Scenario Details").WebButton("Back To Scenario List Link").Click
End If
