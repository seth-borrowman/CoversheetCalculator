Attribute VB_Name = "Module1"
Sub Duplex_CSV():

    Application.ScreenUpdating = False 'Freeze screen
    
    'See if duplexes are dry or wet
    Dim drywet As String
    If Sheets("Wet Duplex").Range("B2").Value <> "" And Sheets("Dry Duplex").Range("B2").Value = "" Then
        drywet = "wet"
    ElseIf Sheets("Dry Duplex").Range("B2").Value <> "" And Sheets("Wet Duplex").Range("B2").Value = "" Then
        drywet = "dry"
    Else
        drywet = "unknown"
    End If
    
    'Count number of duplexes
    Dim num As Integer
    Worksheets("SSRS reorg - Duplexes").Visible = True
    Sheets("SSRS reorg - Duplexes").Select
    num = 0
    For i = 3 To 50
        If Cells(i, 1) <> 0 Then
            num = num + 1
        End If
        Next i
    Worksheets("SSRS reorg - Duplexes").Visible = False
    
    'Check if Janus CSV is hidden, if so then unhide
    If Worksheets("Janus CSV").Visible = False Then
        Worksheets("Janus CSV").Visible = xlSheetVisible
        End If
        
    'Go to CSV
    Sheets("Janus CSV").Select
    
    'Delete every row with a 0 or error resuspension volume
    Dim x As Integer
    x = 1
    While x < 96
        If IsError(Cells(x, 3)) Then
            Rows(x).EntireRow.Delete
            End If
        If IsError(Cells(x, 3)) = False Then
            x = x + 1
            End If
        Wend

    'Save CSV
    Dim path As String
    Dim CurrentName As String
    Dim NameLen As Integer
    Dim NewName As String
    Dim user As String
    
    path = Application.ActiveWorkbook.path 'Get file path
    CurrentName = Application.ActiveWorkbook.FullName 'Get file name
    NameLen = Len(CurrentName) - Len(path) - 1
    CurrentName = Right(CurrentName, NameLen)
    CurrentName = Left(CurrentName, Len(CurrentName) - 5)
    user = Sheets("Wet Duplex").Range("B2").Value & Sheets("Dry Duplex").Range("B2").Value
    
    'Save file before generating CSV
    If ActiveWorkbook.ReadOnly Then 'Make sure that file is not read-only
        MsgBox ("File is Read-only, please save in a new location.")
        If drywet = "dry" Then
            Sheets("Dry Duplex").Select
        Else
            Sheets("Wet Duplex").Select
        End If
        Worksheets("Janus CSV").Visible = False
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
        End
    ElseIf Left(CurrentName, 3) = "PNF" Then 'Make sure file has been saved correctly
        MsgBox ("Please rename file")
        If drywet = "dry" Then
            Sheets("Dry Duplex").Select
        Else
            Sheets("Wet Duplex").Select
        End If
        Worksheets("Janus CSV").Visible = False
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
        End
    Else 'If file is in right place, save before creating CSV
        ActiveWorkbook.Save
    End If
    
    'Set file name and save
    Application.DisplayAlerts = False 'Hide alerts while saving
    ActiveWorkbook.SaveAs Filename:=path & "\" & CurrentName & "_JanusCSV_" & user & CStr(num) & ".csv", FileFormat:=xlCSV, CreateBackup:=True 'Save CSV automatically to same folder
    'CSV name should take the formfile name and just add "_JanusCSV_user#" onto the end
    
    Application.DisplayAlerts = True 'Turn alerts back on
    Application.ScreenUpdating = True 'Unfreeze screen
    ActiveWorkbook.Close SaveChanges:=False 'Close the file

End Sub
