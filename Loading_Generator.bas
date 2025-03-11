option explicit

Sub GenerateLoadingStatements()
    Dim ws As Worksheet
    Dim rng As Range
    Dim data As Variant
    Dim i As Integer
    Dim statements As Collection
    Dim statement As String

    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets(1)
    
    ' Prompt user to select a range
    On Error Resume Next
    Set rng = Application.InputBox("Select the region:", Type:=8)
    On Error GoTo 0
    
    If rng Is Nothing Then Exit Sub
    
    ' Ensure the selected range includes at least 6 columns
    If rng.Columns.Count < 6 Then
        MsgBox "Please select a range with at least 6 columns.", vbExclamation, "Invalid Selection"
        Exit Sub
    End If
    
    ' Inform the user that the process is running
    Application.StatusBar = "Generating statements, please wait..."
    
    ' Store the selected range into a 2D array
    data = rng.Value
    
    ' Initialize the collection to store statements
    Set statements = New Collection
    
    ' Generate statements for each row
    For i = 1 To UBound(data, 1)
        If data(i, 1) = 1 Then
            statement = data(i, 5) & " / " & data(i, 6) & " / " & data(i, 1) & " PKG / " & data(i, 2) & " K"
        Else
            statement = data(i, 5) & " / " & data(i, 6) & " / " & data(i, 1) & " PKGS / " & data(i, 2) & " K"
        End If
        statements.Add statement
    Next i
    
    ' Create the result string
    Dim result As String
    Dim arr() As String
    ReDim arr(1 To statements.Count)
    For i = 1 To statements.Count
        arr(i) = statements(i)
    Next i
    result = Join(arr, vbCrLf)
    
    ' Write the result to a text file on the desktop
    Dim filePath As String
    filePath = Environ("USERPROFILE") & "\Desktop\GeneratedStatements.txt"
    Dim fileNum As Integer
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    Print #fileNum, result
    Close #fileNum
    
    ' Inform the user
    MsgBox "Statements have been written to " & filePath, vbInformation, "Operation Successful"
        
    ' Open the text file
    Shell "notepad.exe " & filePath, vbNormalFocus
    
    ' Reset the status bar
    Application.StatusBar = False
End Sub
