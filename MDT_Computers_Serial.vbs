'
' This script reads a CSV file and writes the information to an MDT SQL Database
' Required columns in CSV file are 'Name', and 'Serial' or 'SN'
' Optional columns can be added to CSV to assign roles, blank fields will be skipped
' If the role does not exist in the database, the script will prompt to add the role
'
Option Explicit

CONST SQLServer = "10.68.4.12"
CONST SQLInstance = "SQLEXPRESS"
CONST SQLDatabase = "MDT"

If (SQLServer = "") OR (SQLInstance = "") OR (SQLDatabase = "") Then
  msgbox "The SQL Constants in the script are blank. Fill in the SQL Constants and run the script again.", &h51000, "Blank Constants"
  NiceQuit()
End If

Dim Conn
Dim strSQLConn, strLogFile
Dim objFSO, objFile, objLog
Dim bolCheck
Dim intLine

Set Conn = CreateObject("ADODB.Connection")
strSQLConn = "Provider=SQLOLEDB.1; Integrated Security=SSPI; Initial Catalog=" & SQLDatabase & "; Data Source=" & SQLServer & "\" & SQLInstance
Conn.Open strSQLConn

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile(ChooseFile, 1)
strLogFile = objFSO.GetParentFolderName(WScript.ScriptFullName) & "\MDTComputers.Log"
Set objLog = objFSO.OpenTextFile(strLogFile, 8, true)
WriteLog("####################################################################################################")
bolCheck = True
intLine = 0

Do while NOT objFile.AtEndOfStream
  intLine = intLine + 1
  Dim arrStr : arrStr = split(objFile.ReadLine,",")
  If bolCheck Then
    bolCheck = CheckCSVHeaders(arrStr(0), arrStr(1))
  Else
    If CheckCSVLine(arrStr(0), arrStr(1), intLine) Then
      Select Case True
        Case (GetCompID(arrStr(0), arrStr(1)) <> "")
          'Existing Computer
          RefreshComputer arrStr
        Case (GetSer(arrStr(0)) <> "")
          'New Serial Existing Name
		  UpdateComputer "UpdateSer", arrStr, GetSer(arrStr(0))
        Case (GetName(arrStr(1)) <> "")
          'New Name Existing Serial
		  UpdateComputer "UpdateName", arrStr, GetName(arrStr(1))
        Case Else
		  'New Computer
          AddNewComputer arrStr
      End Select
    End If
  End If
Loop

Conn.close

msgbox "Finished Adding Computers.", &h51000, "Finished"
WriteLog("Finished adding computers to the database.")
WriteLog("####################################################################################################")

Set objLog = Nothing
Set objFile = Nothing
Set objFSO = Nothing
Set Conn = Nothing

Function ChooseFile()
  With CreateObject("WScript.Shell")
    Dim tempFolder : Set tempFolder = objFSO.GetSpecialFolder(2)
    Dim tempName : tempName = objFSO.GetTempName() & ".hta"
    Dim path : path = "HKCU\Volatile Environment\MsgResp"
    With tempFolder.CreateTextFile(tempName)
      .Write "<input type=file name=f>" & _
        "<script>f.click();(new ActiveXObject('WScript.Shell'))" & _
        ".RegWrite('HKCU\\Volatile Environment\\MsgResp', f.value);" & _
        "close();</script>"
      .Close
    End With
    .Run tempFolder & "\" & tempName, 1, True
    ChooseFile = .RegRead(path)
    .RegDelete path
    objFSO.DeleteFile tempFolder & "\" & tempName
  End With
End Function

Function CheckCSVHeaders(strName, strSer)
  'Check that CSV is formatted correctly
  If (LCase(strName) <> "name") OR ((LCase(strSer) <> "serial") AND (LCase(strSer) <> "sn")) Then
    msgbox "The CSV file's header row is not formatted correctly. The first column should be 'Name', and the second column should be 'Serial' or 'SN'.", &h51000, "Bad Column Header"
    WriteLog("ERROR: The CSV file's header row is not formatted correctly. Exiting script.")
    NiceQuit()
  End If
  CheckCSVHeaders = False
End Function

Function CheckCSVLine(strName, strSer, intLine)
  'Check CSV line for standard errors
  CheckCSVLine = False
  Select Case True
    Case (strName = "")
      'No Name
      WriteLog("ERROR: Cannot add computer with blank name. Skipping line " & intLine & " in the CSV file.")
    Case (strSer = "")
      'No SerialNumber
	  WriteLog("ERROR: The Serial Number for computer name '" & strName & "' is blank. Skipping line " & intLine & " in the CSV file.")
    Case Else
	  CheckCSVLine = True
  End Select
End Function

Function GetCompID(strName, strSer)
  'Get computer ID from database
  dim strSQL
  strSQL = "SELECT ID FROM ComputerIdentity WHERE Description = '" & strName & "' AND SerialNumber = '" & strSer & "'"
  With CreateObject("ADODB.RecordSet")
    .open strSQL,Conn
    On Error Resume Next
    .movefirst
    On Error Goto 0
    If NOT .eof Then
      GetCompID = .Fields(0)
    End if
  End With
End Function

Function GetSer(strName)
  'Get SerialNumber from database using Computer Name
  Dim strSQL
  strSQL = "SELECT SerialNumber FROM ComputerIdentity WHERE Description = '" & strName & "'"
  With CreateObject("ADODB.RecordSet")
    .open strSQL,Conn
    On Error Resume Next
    .movefirst
    On Error Goto 0
    If not .eof Then
      GetSer = .Fields(0)
    End If
  End With
End Function

Function GetName(strSer)
  'Get computer name from database using SerialNumber
  Dim strSQL
  strSQL = "SELECT Description FROM ComputerIdentity WHERE SerialNumber = '" & strSer & "'"
  With CreateObject("ADODB.RecordSet")
    .open strSQL,Conn
    On Error Resume Next
    .movefirst
    On Error Goto 0
    If not .eof Then
      GetName = .Fields(0)
    End If
  End With
End Function

Sub RefreshComputer(arrStr)
  'Refresh the roles on existing computer
  Dim intID, i
  WriteLog("The computer name '" & arrStr(0) & "' with Serial Number '" & arrStr(1) & "' is already in the database.")
  intID = GetCompID(arrStr(0), arrStr(1))
  'Update Roles Assigned to the Computer
  DeleteCompRoles intID, arrStr(0)
  If UBound(arrStr) > 1 Then
    For i = 2 TO UBound(arrStr)
      AssignCompRole intID, i - 1, arrStr(i), arrStr(0)
    Next
  End If
End Sub

Sub UpdateComputer(strType, arrStr, strOld)
  'Update ComputerName description or SerialNumber in database
  Dim strSQL, intID, i
  Select Case strType
    Case "UpdateName"
      strSQL = "UPDATE ComputerIdentity SET Description = '" & arrStr(0) & "' WHERE SerialNumber = '" & arrStr(1) & "'"
	Case "UpdateSer"
	  strSQL = "UPDATE ComputerIdentity SET SerialNumber = '" & arrStr(1) & "' WHERE Description = '" & arrStr(0) & "'"
  End Select
  Conn.Execute(strSQL)
  Select Case strType
    Case "UpdateName"
	  WriteLog("Updated ComputerName for SerialNumber '" & arrStr(1) & "' from '" & strOld & "' to '" & arrStr(0) & "'.")
	Case "UpdateSer"
	  WriteLog("Updated SerialNumber for ComputerName '" & arrStr(0) & "' from '" & strOld & "' to '" & arrStr(1) & "'.")
  End Select
  intID = GetCompID(arrStr(0), arrStr(1))
  'Update ComputerName and OSDComputerName in settings database
  If strType = "UpdateName" Then
    strSQL = "UPDATE Settings SET ComputerName = '" & arrStr(0) & "', OSDComputerName = '" & arrStr(0) & "' WHERE ID = " & intID
    Conn.Execute(strSQL)
    WriteLog("Updated ComputerName and OSDComputerName in Settings database for ComputerName '" & arrStr(0) & "'." )
  End If
  'Update Roles Assigned to the computer
  DeleteCompRoles intID, arrStr(0)
  If UBound(arrStr) > 1 Then
    For i = 2 TO UBound(arrStr)
      AssignCompRole intID, i - 1, arrStr(i), arrStr(0)
    Next
  End If
End Sub

Sub AddNewComputer(arrStr)
  'Insert New Computer into the database
  Dim strSQL, intID, i
  strSQL = "INSERT INTO ComputerIdentity (SerialNumber, Description) VALUES ('" & arrStr(1) & "', '" & arrStr(0) & "')"
  Conn.Execute(strSQL)
  WriteLog("Added '" & arrStr(0) & "' to the database using Serial Number '" & arrStr(1) & "'.")
  intID = GetCompID(arrStr(0), arrStr(1))
  'Insert ComputerName and OSDComputerName settings into the Settings database
  strSQL = "INSERT INTO Settings (Type, ID, ComputerName, OSDComputerName) VALUES ('C', " & intID & ", '" & arrStr(0) & "', '" & arrStr(0) & "')"
  Conn.Execute(strSQL)
  WriteLog("Inserted ComputerName and OSDComputerName in Settings database for ComputerName '" & arrStr(0) & "'." )
  'Add Roles assigned to the computer
  If UBound(arrStr) > 1 Then
    For i = 2 TO UBound(arrStr)
      AssignCompRole intID, i - 1, arrStr(i), arrStr(0)
    Next
  End If
End Sub

Sub DeleteCompRoles(intID, strName)
  'Delete Computer roles from database
  Dim strSQL
  strSQL = "DELETE FROM Settings_Roles WHERE Type='C' AND ID = " & intID
  Conn.Execute(strSQL)
  WriteLog("Deleted all assigned roles from '" & strName & "'.")
End Sub

Sub AssignCompRole(intID, intSeq, strRole, strName)
  'Get proper role name from database
  If strRole = "" Then Exit Sub
  Dim strSQL, strDBRole
  strSQl = "SELECT Role FROM RoleIdentity WHERE Role = '" & strRole & "'"
  With CreateObject("ADODB.RecordSet")
    .open strSQL,Conn
    On Error Resume Next
    .movefirst
    On Error Goto 0
    If not .eof Then
      strDBRole = .Fields(0)
    End if
  End With
  'Assign role to computer
  If NOT strDBRole = "" Then 
    strSQL = "INSERT INTO Settings_Roles (Type, ID, Sequence, Role) VALUES ('C', " & intID & ", " & intSeq & ", '" & strDBRole & "')"
    Conn.Execute(strSQL)
    WriteLog("Assigned '" & strName & "' to the '" & strDBRole & "' role.")
  Else
    Dim strMsg
	strMsg = "Would you like to add the '" & strRole & "' role to the database?" & VBCRLF & VBCRLF &_
			"Click 'Yes' to add this role to the database, 'No' to continue without assigning this role " &_
			"to the computer, or 'Cancel' to exit the script."
    Select Case msgbox(strMsg, &h51000 + 3, "Unknown Role: " & strRole)
      Case vbYes
	    AddRole intID, intSeq, strRole, strName
      Case vbNo
        WriteLog("ERROR: Unable to find role '" & strRole & "' in the database.")
      Case Else
        WriteLog("ERROR: Script file exiting due to user input.")
        NiceQuit()
    End Select
  End If
End Sub

Sub AddRole(intID, intSeq, strRole, strName)
  'Insert Role into database
  Dim strSQL
  strSQL = "INSERT INTO RoleIdentity (Role) VALUES ('" & strRole & "')"
  Conn.Execute(strSQL)
  WriteLog("Added the role '" & strRole & "' to the database.")
  AssignCompRole intID, intSeq, strRole, strName
End Sub

Sub WriteLog(strLine)
  'Write a line to the log file and prefix with date and time
  objLog.WriteLine("[" & Now & "] " & strLine)
End Sub

Sub NiceQuit()
  'Add a # line and then quit
  WriteLog("####################################################################################################")
  wscript.quit
End Sub