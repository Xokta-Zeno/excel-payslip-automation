Option Explicit

Public Sub ExportPayslipNasional()

    Dim wsData As Worksheet
    Dim wsSlip As Worksheet
    Dim headerRow As Long
    
    Set wsData = ActiveWorkbook.Sheets(1)
    Set wsSlip = ActiveWorkbook.Sheets(2)
    
    headerRow = DetectHeaderRow(wsData)
    
    If headerRow = 0 Then
        MsgBox "Header tidak ditemukan!", vbCritical
        Exit Sub
    End If
    
    Dim colUnit As Long, colOutlet As Long, colName As Long, colEmail As Long
    
    colUnit = GetColumnIndex(wsData, headerRow, Array("Pers.Area Desc", "Unit Bisnis", "Pers Area"))
    colOutlet = GetColumnIndex(wsData, headerRow, Array("Cost Center Text", "Outlet"))
    colName = GetColumnIndex(wsData, headerRow, Array("Name", "Employee Name", "Nama"))
    colEmail = GetColumnIndex(wsData, headerRow, Array("alamat email", "email"))
    
    If colUnit = 0 Or colOutlet = 0 Or colName = 0 Or colEmail = 0 Then
        MsgBox "Header penting tidak ditemukan!", vbCritical
        Exit Sub
    End If
    
    Dim periode As String
    periode = CleanName(wsSlip.Range("A16").Value)
    If periode = "" Then periode = Format(Date, "mmmm yyyy")
    
    Dim lastRow As Long
    lastRow = wsData.Cells(wsData.Rows.Count, colName).End(xlUp).Row
    
    Dim basePath As String
    basePath = ActiveWorkbook.Path & "\EXPORT_PDF\"
    If Dir(basePath, vbDirectory) = "" Then MkDir basePath
    
    Dim emailData As Collection
    Set emailData = New Collection
    
    Dim i As Long
    Dim unitName As String, outletName As String
    Dim empName As String, empEmail As String
    Dim folderUnit As String, folderOutlet As String
    Dim pdfName As String
    
    Application.ScreenUpdating = False
    
    For i = headerRow + 1 To lastRow
        
        unitName = CleanName(wsData.Cells(i, colUnit).Value)
        outletName = CleanName(wsData.Cells(i, colOutlet).Value)
        empName = CleanName(wsData.Cells(i, colName).Value)
        empEmail = Trim(wsData.Cells(i, colEmail).Value)
        
        folderUnit = basePath & unitName & "\"
        If Dir(folderUnit, vbDirectory) = "" Then MkDir folderUnit
        
        folderOutlet = folderUnit & outletName & "\"
        If Dir(folderOutlet, vbDirectory) = "" Then MkDir folderOutlet
        
        wsSlip.Range("A6").Value = empName
        
        pdfName = folderOutlet & empName & " - " & periode & ".pdf"
        pdfName = GetSafeFileName(pdfName)
        
        wsSlip.ExportAsFixedFormat _
            Type:=xlTypePDF, _
            Filename:=pdfName, _
            Quality:=xlQualityStandard
        
        If colEmail > 0 Then
        empEmail = Trim(wsData.Cells(i, colEmail).Value)
Else
    empEmail = ""
End If
        
    Next i
    
    Application.ScreenUpdating = True
    
    ZipPerOutlet basePath
    GenerateEmailDistribution emailData, basePath
    
    MsgBox "EXPORT NASIONAL SELESAI + ZIP PER OUTLET + EMAIL LIST!", vbInformation

End Sub

'=========================
' ZIP PER OUTLET
'=========================
Sub ZipPerOutlet(ByVal basePath As String)

    Dim fso As Object
    Dim unitFolder As Object, outletFolder As Object
    Dim shellApp As Object
    Dim zipPath As String, zipFolder As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set shellApp = CreateObject("Shell.Application")
    
    For Each unitFolder In fso.GetFolder(basePath).SubFolders
        For Each outletFolder In unitFolder.SubFolders
            
            zipPath = outletFolder.Path & "\" & outletFolder.Name & ".zip"
            If Dir(zipPath) <> "" Then Kill zipPath
            
            CreateEmptyZip zipPath
            Application.Wait Now + TimeValue("0:00:01")
            
            Set zipFolder = shellApp.Namespace(zipPath)
            
            If Not zipFolder Is Nothing Then
                zipFolder.CopyHere shellApp.Namespace(outletFolder.Path).Items
                Application.Wait Now + TimeValue("0:00:02")
            End If
            
        Next outletFolder
    Next unitFolder

End Sub

'=========================
' GENERATE EMAIL FILE
'=========================
Sub GenerateEmailDistribution(ByVal emailData As Collection, ByVal basePath As String)

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim i As Long
    
    Set wb = Workbooks.Add
    Set ws = wb.Sheets(1)
    
    ws.Range("A1:F1").Value = Array("Nama", "Email", "Unit Bisnis", "Outlet", "Periode", "File PDF")
    
    For i = 1 To emailData.Count
        ws.Cells(i + 1, 1).Resize(1, 6).Value = emailData(i)
    Next i
    
    Dim savePath As String
    savePath = basePath & "EMAIL_DISTRIBUTION.xlsx"
    
    Application.DisplayAlerts = False
    wb.SaveAs savePath, FileFormat:=xlOpenXMLWorkbook
    wb.Close False
    Application.DisplayAlerts = True

End Sub

'=========================
' PREVENT DUPLICATE FILE
'=========================
Function GetSafeFileName(ByVal fullPath As String) As String
    
    Dim fso As Object
    Dim baseName As String, extension As String
    Dim folderPath As String
    Dim counter As Long
    Dim newPath As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FileExists(fullPath) = False Then
        GetSafeFileName = fullPath
        Exit Function
    End If
    
    folderPath = fso.GetParentFolderName(fullPath)
    baseName = fso.GetBaseName(fullPath)
    extension = "." & fso.GetExtensionName(fullPath)
    
    counter = 2
    
    Do
        newPath = folderPath & "\" & baseName & " (" & counter & ")" & extension
        counter = counter + 1
    Loop While fso.FileExists(newPath)
    
    GetSafeFileName = newPath

End Function

'=========================
' CREATE EMPTY ZIP
'=========================
Sub CreateEmptyZip(ByVal zipPath As String)

    Open zipPath For Output As #1
    Print #1, "PK" & Chr(5) & Chr(6) & String(18, Chr(0))
    Close #1

End Sub

'=========================
' AUTO HEADER DETECT
'=========================
Function DetectHeaderRow(ByVal ws As Worksheet) As Long
    
    Dim r As Long, c As Long, lastCol As Long
    
    For r = 1 To 20
        lastCol = ws.Cells(r, ws.Columns.Count).End(xlToLeft).Column
        For c = 1 To lastCol
            If InStr(1, LCase(ws.Cells(r, c).Value), "name") > 0 Then
                DetectHeaderRow = r
                Exit Function
            End If
        Next c
    Next r
    
    DetectHeaderRow = 0

End Function

'=========================
' FLEXIBLE HEADER SEARCH
'=========================
Function GetColumnIndex(ByVal ws As Worksheet, _
                        ByVal headerRow As Long, _
                        ByVal possibleNames As Variant) As Long
    
    Dim c As Long, lastCol As Long
    Dim cellValue As String
    Dim h As Variant
    
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    
    For c = 1 To lastCol
        
        cellValue = ws.Cells(headerRow, c).Value
        cellValue = Trim(LCase(Replace(cellValue, Chr(160), "")))
        
        For Each h In possibleNames
            If cellValue Like "*" & LCase(h) & "*" Then
                GetColumnIndex = c
                Exit Function
            End If
        Next h
        
    Next c
    
    GetColumnIndex = 0
    
End Function

'=========================
' CLEAN FILE NAME
'=========================
Function CleanName(ByVal txt As String) As String
    
    txt = Replace(txt, "\", "")
    txt = Replace(txt, "/", "")
    txt = Replace(txt, ":", "")
    txt = Replace(txt, "*", "")
    txt = Replace(txt, "?", "")
    txt = Replace(txt, """", "")
    txt = Replace(txt, "<", "")
    txt = Replace(txt, ">", "")
    txt = Replace(txt, "|", "")
    
    CleanName = Trim(txt)

End Function
