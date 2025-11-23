Attribute VB_Name = "OptimumUpgrade_Automation"
Option Explicit

' =============================
'  Test Case & Verification Automation
' =============================

Private Function NextTestCaseID(ws As Worksheet) As String
    ' Finds the highest TC-### in column A and returns the next ID.
    Dim lastRow As Long
    Dim i As Long
    Dim maxNum As Long
    Dim curID As String
    Dim numPart As Long

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    maxNum = 0

    For i = 2 To lastRow
        curID = CStr(ws.Cells(i, "A").Value)
        If Len(curID) > 3 And Left(curID, 3) = "TC-" Then
            On Error Resume Next
            numPart = CLng(Mid(curID, 4))
            On Error GoTo 0
            If numPart > maxNum Then maxNum = numPart
        End If
    Next i

    maxNum = maxNum + 1
    NextTestCaseID = "TC-" & Format(maxNum, "000")
End Function

Public Sub GenerateTestCasesFromRequirements()
    Dim wsReq As Worksheet
    Dim wsTC As Worksheet
    Dim lastReqRow As Long
    Dim lastTcRow As Long
    Dim r As Long
    Dim reqID As String
    Dim reqText As String
    Dim method As String
    Dim tcID As String
    Dim existingTCID As String

    On Error GoTo ErrHandler

    Set wsReq = ThisWorkbook.Worksheets("Requirements")
    Set wsTC = ThisWorkbook.Worksheets("TestCases")

    lastReqRow = wsReq.Cells(wsReq.Rows.Count, "A").End(xlUp).Row

    For r = 2 To lastReqRow
        reqID = Trim(CStr(wsReq.Cells(r, "A").Value))
        reqText = Trim(CStr(wsReq.Cells(r, "B").Value))
        method = LCase(Trim(CStr(wsReq.Cells(r, "E").Value))) ' Verification Method
        existingTCID = Trim(CStr(wsReq.Cells(r, "F").Value))  ' Test Case ID

        ' Only generate for requirements that need testing and don't already have a TC
        If reqID <> "" Then
            If existingTCID = "" Then
                If (InStr(method, "test") > 0) Or (InStr(method, "demonstration") > 0) Then
                    ' Get next TC ID
                    tcID = NextTestCaseID(wsTC)

                    ' Write TC ID back into Requirements
                    wsReq.Cells(r, "F").Value = tcID

                    ' Append a new row in TestCases
                    lastTcRow = wsTC.Cells(wsTC.Rows.Count, "A").End(xlUp).Row + 1
                    wsTC.Cells(lastTcRow, "A").Value = tcID                                ' Test Case ID
                    wsTC.Cells(lastTcRow, "B").Value = "Auto-gen for " & reqID            ' Title
                    wsTC.Cells(lastTcRow, "C").Value = reqID                               ' Related Req ID(s)
                    wsTC.Cells(lastTcRow, "D").Value = "Verify: " & reqText               ' Description
                    wsTC.Cells(lastTcRow, "E").Value = "System Test"                       ' Type (default)
                    wsTC.Cells(lastTcRow, "F").Value = "Production-representative config"  ' Configuration
                    wsTC.Cells(lastTcRow, "G").Value = "Define detailed steps..."          ' Steps
                    wsTC.Cells(lastTcRow, "H").Value = "Requirement " & reqID & " met."    ' Expected Result
                    wsTC.Cells(lastTcRow, "I").Value = "Planned"                           ' Status
                End If
            End If
        End If
    Next r

    MsgBox "Test case generation complete.", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Error in GenerateTestCasesFromRequirements: " & Err.Description, vbCritical
End Sub

Public Sub BuildVerificationSummary()
    Dim wsReq As Worksheet
    Dim wsVer As Worksheet
    Dim lastReqRow As Long
    Dim r As Long
    Dim outRow As Long

    On Error GoTo ErrHandler

    Set wsReq = ThisWorkbook.Worksheets("Requirements")

    On Error Resume Next
    Set wsVer = ThisWorkbook.Worksheets("VerificationSummary")
    On Error GoTo 0

    If wsVer Is Nothing Then
        Set wsVer = ThisWorkbook.Worksheets.Add(After:=wsReq)
        wsVer.Name = "VerificationSummary"
    Else
        wsVer.Cells.Clear
    End If

    ' Header
    wsVer.Range("A1:E1").Value = Array("Req ID", "Requirement", "Verification Method", "Test Case ID", "Status")

    lastReqRow = wsReq.Cells(wsReq.Rows.Count, "A").End(xlUp).Row
    outRow = 2

    For r = 2 To lastReqRow
        If Trim(CStr(wsReq.Cells(r, "A").Value)) <> "" Then
            wsVer.Cells(outRow, "A").Value = wsReq.Cells(r, "A").Value ' Req ID
            wsVer.Cells(outRow, "B").Value = wsReq.Cells(r, "B").Value ' Requirement Text
            wsVer.Cells(outRow, "C").Value = wsReq.Cells(r, "E").Value ' Verification Method
            wsVer.Cells(outRow, "D").Value = wsReq.Cells(r, "F").Value ' Test Case ID
            wsVer.Cells(outRow, "E").Value = wsReq.Cells(r, "G").Value ' Status
            outRow = outRow + 1
        End If
    Next r

    wsVer.Columns("A:E").AutoFit

    MsgBox "VerificationSummary sheet rebuilt.", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Error in BuildVerificationSummary: " & Err.Description, vbCritical
End Sub

Public Sub ExportVerificationSummaryCSV()
    Dim wsVer As Worksheet
    Dim fName As Variant

    On Error GoTo ErrHandler

    Set wsVer = ThisWorkbook.Worksheets("VerificationSummary")

    fName = Application.GetSaveAsFilename( _
        InitialFileName:="VerificationSummary.csv", _
        FileFilter:="CSV Files (*.csv), *.csv")

    If fName = False Then
        Exit Sub
    End If

    wsVer.Copy ' copies to a new temporary workbook
    With ActiveWorkbook
        .SaveAs Filename:=fName, FileFormat:=xlCSV, CreateBackup:=False
        .Close SaveChanges:=False
    End With

    MsgBox "VerificationSummary exported to: " & fName, vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Error in ExportVerificationSummaryCSV: " & Err.Description, vbCritical
End Sub

' =============================
'  Project Info Automation
' =============================

' Minimal silent install helper called by Workbook_Open
Public Sub Install_Automation()
    On Error Resume Next
    Call CreateProjectInfoIfMissing
    ' Add other silent initialization steps here if required
    Call CreateControlPanelIfMissing
End Sub

Private Sub CreateProjectInfoIfMissing()
    Dim ws As Worksheet
    Dim tbl As Variant
    Dim i As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("ProjectInfo")
    On Error GoTo 0

    If Not ws Is Nothing Then Exit Sub

    Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
    ws.Name = "ProjectInfo"

    ws.Range("A1").Value = "Field"
    ws.Range("B1").Value = "Value"

    tbl = Array( _
        "ProjectName", _
        "ProjectAcronym", _
        "Customer", _
        "ContractNumber", _
        "ProgramID", _
        "PMName", _
        "SELead", _
        "StartDate", _
        "EndDate" _
    )

    For i = LBound(tbl) To UBound(tbl)
        ws.Cells(i + 2, "A").Value = tbl(i)
        ws.Cells(i + 2, "B").Value = ""
    Next i

    ws.Columns("A:B").AutoFit
End Sub

Private Sub CreateControlPanelIfMissing()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("ControlPanel")
    On Error GoTo 0

    If Not ws Is Nothing Then Exit Sub

    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = "ControlPanel"

    ws.Range("A1").Value = "Automation Control Panel"
    ws.Range("A2").Value = "Run Install_Automation to configure project defaults." 
    ws.Range("A1:A2").Font.Bold = True
    ws.Columns("A").ColumnWidth = 50
End Sub

Public Sub EnsureProjectInfoSheet()
    Dim ws As Worksheet
    Dim tbl As Variant
    Dim i As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("ProjectInfo")
    On Error GoTo 0

    If Not ws Is Nothing Then
        MsgBox "ProjectInfo sheet already exists.", vbInformation
        Exit Sub
    End If

    Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
    ws.Name = "ProjectInfo"

    ws.Range("A1").Value = "Field"
    ws.Range("B1").Value = "Value"

    tbl = Array( _
        "ProjectName", _
        "ProjectAcronym", _
        "Customer", _
        "ContractNumber", _
        "ProgramID", _
        "PMName", _
        "SELead", _
        "StartDate", _
        "EndDate" _
    )

    For i = LBound(tbl) To UBound(tbl)
        ws.Cells(i + 2, "A").Value = tbl(i)
        ws.Cells(i + 2, "B").Value = ""  ' to be filled by user
    Next i

    ws.Columns("A:B").AutoFit

    MsgBox "ProjectInfo sheet created. Please fill in the values.", vbInformation
End Sub

Public Sub UpdateSheetHeadersFromProjectInfo()
    Dim wsInfo As Worksheet
    Dim ws As Worksheet
    Dim projName As String
    Dim projAcr As String
    Dim customer As String
    Dim contractNo As String
    Dim progID As String
    Dim headerTitle As String
    Dim headerSub As String

    On Error GoTo ErrHandler

    Set wsInfo = ThisWorkbook.Worksheets("ProjectInfo")

    ' Read values from the ProjectInfo table
    projName = GetProjectField(wsInfo, "ProjectName")
    projAcr = GetProjectField(wsInfo, "ProjectAcronym")
    customer = GetProjectField(wsInfo, "Customer")
    contractNo = GetProjectField(wsInfo, "ContractNumber")
    progID = GetProjectField(wsInfo, "ProgramID")

    ' Loop through all sheets and update A1/A2 (except ProjectInfo itself)
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "ProjectInfo" Then
            headerTitle = projName
            If projAcr <> "" Then
                headerTitle = projName & " (" & projAcr & ")"
            End If

            headerSub = ""
            If progID <> "" Then headerSub = headerSub & "Program ID: " & progID
            If customer <> "" Then
                If headerSub <> "" Then headerSub = headerSub & " | "
                headerSub = headerSub & "Customer: " & customer
            End If
            If contractNo <> "" Then
                If headerSub <> "" Then headerSub = headerSub & " | "
                headerSub = headerSub & "Contract: " & contractNo
            End If

            ws.Range("A1").Value = headerTitle
            ws.Range("A2").Value = headerSub
            ws.Rows("1:2").Font.Bold = True
        End If
    Next ws

    MsgBox "Sheet headers updated from ProjectInfo.", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Error in UpdateSheetHeadersFromProjectInfo: " & Err.Description, vbCritical
End Sub

Private Function GetProjectField(wsInfo As Worksheet, fieldName As String) As String
    Dim lastRow As Long
    Dim r As Long

    lastRow = wsInfo.Cells(wsInfo.Rows.Count, "A").End(xlUp).Row
    For r = 2 To lastRow
        If LCase(Trim(CStr(wsInfo.Cells(r, "A").Value))) = LCase(Trim(fieldName)) Then
            GetProjectField = CStr(wsInfo.Cells(r, "B").Value)
            Exit Function
        End If
    Next r

    GetProjectField = ""
End Function
