
'###########################################################################
'Script Name: UFT_Get_All_Objects_Names-GitHub_Demo
'Script Path: C:\Users\Donald Ingreson\Documents\GitHub\UFT_GITHUB_DEMO\UFT_Get_All_Objects_Names-GitHub_Demo
'Description: Script gets list of all objects with names and stores in Global DataTable
'JIRA JOMIS Type: N/A
'JIRA JOMIS ID number: N/A
'Add-ins to use: Web
'Run Settings Iterations: Run one iteration only
'Pre-conditions: AUT Web page is opened after UFT
'Post-conditions: N/A
'Developed by: Don Ingerson
'Date of Creation: 10/6/2021
'Date of last Modification: 
'Reason for Modification: N/A
'Latest version: Demo
'###########################################################################

Option Explicit

Dim oWebPageObjects
Dim intNoOfObjects
Dim ii
Dim jj
Dim nWebBrowserCount

'Close and open Chrome Browsers
'SystemUtil.CloseProcessByName "chrome.exe"

'## https://nolijconsultingllc.sharepoint.com/sites/DHMS%20PEO%20TESS/JOMIS/wiki/Home.aspx

nWebBrowserCount = fn_Browser_Tabs_Count()
Print "nWebBrowserCount = " & nWebBrowserCount

If (nWebBrowserCount = 0) Then
	msgbox("There is no IE Browser Open - Test Terminated")
	ExitTest
End If

If (nWebBrowserCount > 1) Then
	msgbox("More than one IE Browser/Tab open - Test Terminated")
	ExitTest
End If

'Get all Child Objects
Set oWebPageObjects = Browser("name:=.*").Page("title:=.*").ChildObjects()
intNoOfObjects = oWebPageObjects.Count
Print "intNoOfObjects = " & intNoOfObjects

'Add Columns to the Global DataTable
DataTable.GlobalSheet.AddParameter "ObjectType", ""
DataTable.GlobalSheet.AddParameter "ObjectName", ""

'Get all object types and names and write to Global DataTable
For ii = 0 to intNoOfObjects - 1
    jj = ii+1
	DataTable.SetCurrentRow(jj)
	DataTable.Value("ObjectType") = oWebPageObjects(ii).GetROProperty("micclass")
	Print "jj = " & jj
'	oWebPageObjects(ii).Highlight
	Print "Type = " & oWebPageObjects(ii).GetROProperty("micclass")
	DataTable.Value("ObjectName") = oWebPageObjects(ii).GetROProperty("name")
	Print "ObjectName = " & oWebPageObjects(ii).GetROProperty("name")
	
Next    

'Release the memory
Set oWebPageObjects = Nothing

'SystemUtil.CloseProcessByName "iexplore.exe"
' SystemUtil.CloseProcessByName "chrome.exe"

Function fn_Browser_Tabs_Count()
	Dim oDesc
	Dim Obj2
	Dim nCnt	
	
	Set oDesc = Description.Create()
	oDesc("micclass").Value = "Browser"
	
	Set Obj2 = DeskTop.ChildObjects(oDesc)
	
	nCnt = Obj2.Count	

	fn_Browser_Tabs_Count = nCnt
	
End Function
