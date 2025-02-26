Private Sub ClearBtn_Click()
    Dim OpenWS As Worksheet: Set OpenWS = ThisWorkbook.Worksheets("Final")
    OpenWS.Range("A6:L65536").ClearContents
    OpenWS.Range("O6:O65536").ClearContents
    ClearBtn.Font.Size = 12
End Sub
Private Sub Populate()
    'TODO
    'Ask for file
    'Get all on hand
    'Separate by customer
    'Generate Delivery Lists
    
    'Column links
    Dim CopyDict As New Scripting.Dictionary
    CopyDict.Add "Columns", New Scripting.Dictionary
    CopyDict("Columns").Add "D", "C" 'Customer Name
    CopyDict("Columns").Add "B", "E" 'Customer Code
    CopyDict("Columns").Add "H", "F" 'DNPH Part Number
    CopyDict("Columns").Add "F", "G" 'PO Number
    CopyDict("Columns").Add "O", "H" 'Description
    CopyDict("Columns").Add "M", "I" 'Qty
    CopyDict("Columns").Add "AM", "J" 'Curr
    CopyDict("Columns").Add "AN", "K" 'Price
    CopyDict("Columns").Add "AO", "L" 'Rate
    
    OpenFile = Application.GetOpenFilename(Title:="Browse", FileFilter:="Excel Files (*.xls*),*xls*")
    If OpenFile = False Then
        MsgBox "No file selected", vbExclamation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.AskToUpdateLinks = False
    Application.DisplayAlerts = False
    ActiveWindow.View = xlNormalView
    Application.Calculation = xlCalculationManual
    
    'Ask for file
    Dim OpenBook As Workbook: Set OpenBook = Application.Workbooks.Open(OpenFile)
    Dim OpenSheet As Worksheet
    
    'Get sheetname with validation
    Do While True
        sheetName = InputBox("Please enter the name of the sheet: ")
            
        If sheetName = "" Then
            MsgBox "No sheet selected", vbRetryCancel
            If Response = vbCancel Then
                Exit Sub
            End If
        End If
    
        On Error Resume Next
        Set OpenSheet = OpenBook.Sheets(sheetName)
        
        If OpenSheet Is Nothing Then
            Response = MsgBox("Worksheet not found", vbRetryCancel)
            If Response = vbCancel Then
                Exit Sub
            End If
        Else
            Exit Do
        End If
    Loop
    
    Debug.Print ("Start Time: " & Time)
    Call DeleteSheets
    
    Dim CountsDict As New Scripting.Dictionary
    For I = 3 To 65536
        Customer = OpenSheet.Cells(I, 4).Value
        
        If Customer = "" Then
            Exit For
        End If
        
        'Get all on hand
        If OpenSheet.Range("AV" & I).Value = "ON HAND" Then
            'Separate by customer
            If Not CountsDict.Exists(Customer) Then
                CountsDict.Add Customer, 0
                Call CreateNewSheet(CStr(Customer))
            End If
            
            Count = CountsDict(Customer) + 1
            CountsDict(Customer) = Count
            
            Dim ThisSheet As Worksheet: Set ThisSheet = ThisWorkbook.Sheets(Customer)
            Row = Count + 5
            ThisSheet.Range("B" & Row).Value = Count
            If OpenSheet.Range("BC" & I).Value = "N1" Then
                ThisSheet.Range("D" & Row).Value = 1
            ElseIf OpenSheet.Range("BC" & I).Value = "M1" Then
                ThisSheet.Range("D" & Row).Value = 7
            End If
            For Each col In CopyDict("Columns").Keys
                ThisSheet.Range(CopyDict("Columns")(col) & Row).Value = OpenSheet.Range(col & I).Value
            Next col
        End If
    Next I
    
    'Generate Delivery Lists
    For Each Key In CountsDict.Keys
        GenerateFile (CStr(Key))
    Next Key
    Debug.Print ("End Time: " & Time)
    
    Application.Calculate
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    OpenBook.Close(True)
    ThisWorkbook.Save
    MsgBox "Finished"
End Sub

Private Sub CreateNewSheet(sheetName As String)
    Dim templateWS As Worksheet: Set templateWS = ThisWorkbook.Worksheets("Template")
    templateWS.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Name = sheetName
End Sub

Private Sub ClearSheets(sheetName As String)
    Dim OpenWS As Worksheet: Set OpenWS = ThisWorkbook.Worksheets(sheetName)
    OpenWS.Range("A6:L65536").ClearContents
    OpenWS.Range("O6:O65536").ClearContents
End Sub

Private Sub DeleteSheets()
    Application.DisplayAlerts = False
    For Each Sheet In ThisWorkbook.Sheets
        If Not (Sheet.Name = "Final" Or Sheet.Name = "Template") Then
            Sheet.Delete
        End If
    Next Sheet
    Application.DisplayAlerts = True
End Sub

Private Sub GenerateFile(sheetName As String)
    Application.DisplayAlerts = False
    Dim NewBook As Workbook: Set NewBook = Workbooks.Add
    Dim FromSheet As Worksheet: Set FromSheet = ThisWorkbook.Worksheets(sheetName)
    FromSheet.Range("B2").Value = "Customer Name: " & sheetName
    FromSheet.Range("B3").Value = "Delivery Date: " & Format(Date, "mmmm d, yyyy")
    FromSheet.Copy Before:=NewBook.Sheets(1)
    FromSheet.Delete
    FilePath = ThisWorkbook.Path & "\Output\" & sheetName & ".xlsx"
    If Dir(FilePath) <> "" Then
        Kill (FilePath)
    End If
    NewBook.SaveAs FilePath
    NewBook.Close (True)
    Application.DisplayAlerts = True
End Sub

Private Sub PopulateBtn_Click()
    Call Populate
End Sub

