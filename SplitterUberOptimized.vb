Sub Splitter()

    Dim FldrPicker As FileDialog
    Dim myFolder As String
    Dim fileName As String
    Dim values() As Variant
    Dim valueA As String
    Dim valueB As String
    Dim startTime As Double
    Dim endTime As Double
    Dim totalTime As Double
    Dim i As Long
    Dim j As Long

    'Have User Select Folder to Save to with Dialog Box
    Set FldrPicker = Application.FileDialog(msoFileDialogFolderPicker)
    With FldrPicker
        .Title = "Select a Target Folder"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Sub 'Check if user clicked cancel button
        myFolder = .SelectedItems(1) & "\"
    End With
    
    'Prompt user for fileName
    fileName = InputBox("Enter the filename prefix: ")
    
    'Remove '_' if included
    If Right$(fileName, 1) = "_" Then
        fileName = Left(fileName, Len(fileName) - 1)
    End If
    
    Application.ScreenUpdating = False
    startTime = Timer 'start time
    'Get the last row of data
    lastRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    ReDim values(1 To lastRow, 1 To 2)
    'Loop through the data in columns A and B
    For i = 1 To lastRow
        'Get the value in column A
        valueA = UCase(ActiveSheet.Cells(i, 1).Value)
        'Get the value in column B
        valueB = ActiveSheet.Cells(i, 2).Value
        'store the value in array
        values(i, 1) = valueA
        values(i, 2) = valueB
    Next i
    'Loop through the array and save the values in column A as separate CSV files
    For i = 1 To lastRow
        If values(i, 2) <> 0 And values(i, 1) <> "Acct" Then
            'Create a new CSV file
            Set NewFile = CreateObject("Scripting.FileSystemObject").CreateTextFile(myFolder & fileName & "_" & CStr(values(i, 2)) & ".csv")
            'Write the values in column A to the file
            NewFile.Write values(i, 1)
            'Close the file
            NewFile.Close
	Next Key
		
		endTime = Now()
		totalTime = DateDiff("s", startTime, endTime)
		Call Shell("explorer.exe" & " " & myFolder, vbNormalFocus)
		MsgBox "Splitter Finished " & myFolder & vbNewLine & " Total time: " & totalTime & " seconds"
		
	End Sub