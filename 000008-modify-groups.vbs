''	---------------------------------------------------------------------------
''
''
''	SCRIPT
''		modify-groups.vbs
''
''
''	SCRIPT_ID
''		8
''
'' 
''	DESCRIPTION
''		Read from a Excel file the columns and add of removes accounts from groups
''
''
''	VERSION
''		01	2015-04-22	Modified to use an Excel file as input file.
''		02	????-??-??	Modifications
''		03	2014-01-31	First version
''
'' 
''	FUNCTIONS AND SUBS
''		Function GetWordsheetIdFromName
''		Function GetScriptName
''		Function GetScriptPath
''		Function GetIdFromHeaderColumn
'' 		Function SelectWorkSheet
''		Sub ScriptUsage
''		Sub ScriptInit
''		Sub ScriptRun
''		Sub ScriptDone
'' 
''
''	---------------------------------------------------------------------------

Option Explicit

''	---------------------------------------------------------------------------

Dim		gobjExcel
Dim		gobjSheet
Dim		gintWorksheetCount
Dim		gobjFso
Dim		gstrPathExcel

''	---------------------------------------------------------------------------

Function GetScriptName()
	''
	''	Returns the script name
	''	Removes script versioning (script-00)
	''	Removes .vbs extension
	''
	Dim	strReturn

	strReturn = WScript.ScriptName
	strReturn = Replace(strReturn, ".vbs", "")			'' Set the script name (without .vbs extention)

	If Mid(strReturn, Len(strReturn) - 2, 1) = "-" Then
		strReturn = Left(strReturn, Len(strReturn) - 3)
	End If
	GetScriptName = strReturn
End Function '' GetScriptName()



Function EncloseWithDQ(ByVal s)
	''
	''	Returns an enclosed string s with double quotes around it.
	''	Check for exising quotes before adding adding.
	''
	''	s > "s"
	''
	
	If Left(s, 1) <> Chr(34) Then
		s = Chr(34) & s
	End If
	
	If Right(s, 1) <> Chr(34) Then
		s = s & Chr(34)
	End If

	EncloseWithDQ = s
End Function '' of Function EncloseWithDQ



Function GetScriptPath()
	'==
	'==	Returns the path where the script is located.
	'==
	'==	Output:
	'==		A string with the path where the script is run from.
	'==
	'==		drive:\folder\folder\
	'==
	Dim sScriptPath
	Dim sScriptName

	sScriptPath = WScript.ScriptFullName
	sScriptName = WScript.ScriptName

	GetScriptPath = Left(sScriptPath, Len(sScriptPath) - Len(sScriptName))
End Function '' GetScriptPath



Function DsQueryGetDn(ByVal strRootDse, ByVal strCn)
	''
	''	Use the DSQUERY.EXE command to find a DN of a CN in a specific AD set by strRootDse
	''
	''		strRootDse: DC=prod,DC=ns,DC=nl
	''		strCn: 		ZZZ_NAME_OF_GROUP
	''
	''		Returns: 	The DN of blank if not found.
	''
	
	Dim		c			''	Command
	Dim		r			''	Result
	Dim		objShell
	Dim		objExec
	Dim		strOutput
	
	If InStr(strCn, "CN=") > 0 Then
		'' When the strCN already contains a Distinguished Name (DN), result = strCn
		r = EncloseWithDQ(strCn)
	Else
		'' No, we must search for the DN based on the CN
	
		c = "dsquery.exe "
		c = c & "* "
		c = c & strRootDse & " "
		c = c & "-filter (CN=" & strCn & ")"

		Set objShell = CreateObject("WScript.Shell")
		Set objExec = objShell.Exec(c)
		
		Do
			strOutput = objExec.Stdout.ReadLine()
		Loop While Not objExec.Stdout.atEndOfStream

		Set objExec = Nothing
		Set objShell = Nothing
		If Len(strOutput) > 0 Then
			r = EncloseWithDQ(strOutput)  '' BEWARE: r contains now " around the string, see "CN=name,OU=name,DC=domain,DC=nl"
		Else
			WScript.Echo "ERROR Could not find the Distinguished Name for " & strCn & " in " & strRootDse
			r = ""
		End If
	End If
	DsQueryGetDn = r
End Function '' DsQueryGetDn



Function GetWordsheetIdFromName(ByVal strSearchName)
	'
	'	Search for the worksheet ID using strSearchName
	'	Case sensitive
	'
	'	Returns:
	'		0	Not found
	'		<>0	Worksheet ID associated with strSearchName
	'
	
	Dim 	currentWorkSheet
	Dim		r
	Dim		counter
	
	r = 0
	
	For counter = 1 to gintWorksheetCount
		Set currentWorkSheet = gobjExcel.ActiveWorkbook.Worksheets(counter)
		'WScript.Echo "Current Worksheet name: " & currentWorkSheet.Name & ", search for: " & strSearchName
		If strSearchName = currentWorkSheet.Name Then
			r = counter
			Exit For
		End If
		Set currentWorkSheet = Nothing
	Next
	GetWordsheetIdFromName = r
End Function



Function GetIdFromHeaderColumn(ByVal objExcel, ByVal intWorksheetId, ByVal strColumnName)
	''
	''	Returns the column number where the strColumnName is the header column name
	''
	Dim		c
	Dim		objWorksheet
	Dim		intMaxColumn
	Dim		objCells
	Dim		strCellContents
	Dim		r
	Dim		blnFound
	
	Set objWorksheet = objExcel.ActiveWorkbook.Worksheets(intWorksheetId)
	
	intMaxColumn = objWorksheet.UsedRange.Columns.Count
	
	Set objCells = objWorksheet.Cells
	blnFound = False
	
	For c = 1 To intMaxColumn
		strCellContents = objCells(1, c).Value
		'WScript.Echo strCellContents
		If strCellContents = strColumnName Then
			r = c
			blnFound = True
		End If
	Next
	
	If blnFound = True Then	
		GetIdFromHeaderColumn = r
	Else
		GetIdFromHeaderColumn = 0
		WScript.Echo "GetIdFromColumnName(" & strColumnName & ") ERROR could not find it!"
		
		gobjExcel.Quit
		Set gobjExcel = Nothing

		WScript.Quit(0)
	End If
	
	Set objCells = Nothing
	Set objWorksheet = Nothing
End Function



Function SelectWorkSheet(ByRef objExcel, ByVal strWorksheetName)
	'' 
	''	Sets the current worksheet by name
	''
	
	Dim		counter
	Dim		currentWorkSheet
	Dim		blnFound
	Dim		r
	
	blnFound = False
	r = 0

	For counter = 1 to objExcel.Worksheets.Count
		Set currentWorkSheet = gobjExcel.ActiveWorkbook.Worksheets(counter)
		'WScript.Echo "Current Worksheet name: " & currentWorkSheet.Name & ", searching for: " & strSearchName
		If strWorksheetName = currentWorkSheet.Name Then
			blnFound = True
			objExcel.ActiveWorkbook.Sheets(counter).Select
			r = counter
			Exit For
		End If
		Set currentWorkSheet = Nothing
	Next
	
	If blnFound = False Then
		WScript.Echo "SelectWorkSheet(" & strWorksheetName & "): Could not find worksheet in " & gstrPathExcel & ", stopping..."

		gobjExcel.Quit
		Set gobjExcel = Nothing
		
		WScript.Quit(0)
	Else
		WScript.Echo "Current worksheet is now: " & strWorksheetName
	End If
	
	SelectWorkSheet = r
End Function '' of Sub SelectWorkSheet

''	---------------------------------------------------------------------------

Sub ModifyGroup(ByVal strGroupDn, ByVal strAction, ByVal strAccountDn)
	Dim		c			'' Command Line 
	Dim		r			'' Result
	Dim		strOption	'' Use option for DSMOD.EXE
	Dim		objShell
	
	WScript.Echo "ModifyGroup():"
	
	If UCase(strAction) = "ADD" Then
		strOption = "-addmbr"
	ElseIf UCase(strAction) = "DEL" Then
		strOption =  "-rmmbr"
	Else
		Wscript.Echo "DsGroupModifyMember(): Wrong action code for strAction specified: " & strAction
		r = -9
		Exit Sub
	End If
	
	c = "dsmod.exe group " & strGroupDn & " " & strOption & " " & strAccountDn
	WScript.Echo c
	
	Set objShell = CreateObject("WScript.Shell")
	
	c = "cmd.exe /c " & c
	r = objShell.Run(c, 0, True)
	
	Set objShell = Nothing
	If r > 0 Then
		WScript.Echo "  ERROR LEVEL=0x" & Hex(r)
	Else
		WScript.Echo "  OK"
	End If
	
	WScript.Echo
	
End Sub	'' of Sub ModifyGroup



Sub ScriptUsage()
	
	WScript.Echo "Usage: " & GetScriptName() & " /file:<subfolder name>"
	WScript.Echo
	WScript.Echo vbTab & "/file" & vbTab & "Name of Excel to to be processed"
	WScript.Echo 
	WScript.Echo "Excel file needs to be compliant to:"
	WScript.Echo " - Worksheet name: INPUT"
	WScript.Echo " - Column name A: RootDse"
	WScript.Echo " - Column name B: Group"
	WScript.Echo " - Column name C: Action"
	WScript.Echo " - Column name D: Account"
	Wscript.Echo
		
	WScript.Quit(0)
End Sub '' ScriptUsage()



Sub ScriptInit()
	Dim		colNamedArguments
	Dim		intArgumentCount

	
	Set gobjFso = CreateObject("Scripting.FileSystemObject")
	
	Set colNamedArguments = WScript.Arguments.Named
	intArgumentCount = WScript.Arguments.Named.Count
	
	If intArgumentCount <> 1 Then
		Call ScriptUsage()
		Set colNamedArguments = Nothing
	End If
	
	gstrPathExcel = WScript.Arguments.Named("file")
	If InStr(gstrPathExcel, ":\") = 0 Then
		gstrPathExcel = GetScriptPath & gstrPathExcel
	End If
	
	If gobjFso.FileExists(gstrPathExcel) = False Then
		WScript.Echo "ERROR Could not find the file " & gstrPathExcel
		WScript.Quit(0)
	End If
End Sub '' of Sub ScriptInit



Sub ScriptRun
	Const	COL_FIRST	=	1

	Dim		strPathExcel
	Dim		intRow
	Dim		strWorksheetName
	Dim		objWorksheet
	Dim		intWorksheetId
	Dim		intColumnRootDse
	Dim		intColumnGroup
	Dim		intColumnAction
	Dim		intColumnAccount
	Dim		strRootDse
	Dim		strGroup
	Dim		strAction
	Dim		strAccount
	Dim		strGroupDn
	
	'strPathExcel = GetScriptPath & "test.xlsx"
	'gstrPathExcel = GetScriptPath & "input.xlsx"
	
	WScript.Echo "Starting Excel and open " & gstrPathExcel
	
	Set gobjExcel = CreateObject("Excel.Application")
	Set gobjSheet = gobjExcel.Workbooks.Open(gstrPathExcel)
	
	gintWorksheetCount = gobjExcel.Worksheets.Count
	WScript.Echo "Workbook has " & gintWorksheetCount & " sheets"
	
	strWorksheetName = "INPUT"
	intWorksheetId = SelectWorkSheet(gobjExcel, strWorksheetName)
	
	WScript.Echo "Worksheet " & strWorksheetName & " has ID " & intWorksheetId

	intColumnRootDse = GetIdFromHeaderColumn(gobjExcel, intWorksheetId, "RootDse")
	intColumnGroup = GetIdFromHeaderColumn(gobjExcel, intWorksheetId, "Group")
	intColumnAction = GetIdFromHeaderColumn(gobjExcel, intWorksheetId, "Action")
	intColumnAccount = GetIdFromHeaderColumn(gobjExcel, intWorksheetId, "Account")
	
	
	intRow = 2 '' Start on row 2, row 1 is the header row
	
	Do Until gobjExcel.Cells(intRow, 1).Value = ""
		
		strRootDse = Trim(gobjExcel.Cells(intRow, intColumnRootDse).Value)
		strGroup = Trim(gobjExcel.Cells(intRow, intColumnGroup).Value)
		strAction = Trim(gobjExcel.Cells(intRow, intColumnAction).Value)
		strAccount = Trim(gobjExcel.Cells(intRow, intColumnAccount).Value)
		
		'WScript.Echo strRootDse & vbTab & strGroup & " " & strAction & " " & strAccount
		strGroupDn = DsQueryGetDn(strRootDse, strGroup)
		If Len(strGroupDn) > 0  Then
			Call ModifyGroup(strGroupDn, strAction,DsQueryGetDn(strRootDse, strAccount))
		End If
		
		
		intRow = intRow + 1
	Loop '' Do Until gobjExcel.Cells(intRow, 1).Value = ""
	
	
	
	Set gobjSheet = Nothing
	
	gobjExcel.Quit
	Set gobjExcel = Nothing
End Sub '' of Sub ScriptRun



Sub ScriptDone()
	Set gobjFso = Nothing
End Sub '' of Sub ScriptDone



Sub ScriptTest()
	WScript.Echo DsQueryGetDn("DC=prod,DC=ns,DC=nl", "RSA_E0600_ENTERPRISE_DGG")
	WScript.Echo DsQueryGetDn("DC=prod,DC=ns,DC=nl", "CN=RSA_E0600_ENTERPRISE_DGG,OU=Domain Global Groups,OU=Resources,DC=prod,DC=ns,DC=nl")
	
	WScript.Echo EncloseWithDQ("test")
	WScript.Echo EncloseWithDQ(Chr(34) & "test")
	WScript.Echo EncloseWithDQ("test" & Chr(34))
	WScript.Echo EncloseWithDQ(Chr(34) & "test" & Chr(34))
End Sub '' of Sub ScriptTest

''	---------------------------------------------------------------------------

Call ScriptInit()
Call ScriptRun()
'Call ScriptTest()
Call ScriptDone()
WScript.Quit(0)

''	---------------------------------------------------------------------------