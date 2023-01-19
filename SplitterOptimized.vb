Sub Splitter()

    Dim FldrPicker As FileDialog
    Dim myFolder As String
    Dim fileName As String
    Dim values As Object
    Dim valueA As String
    Dim valueB As String
    Dim currentArray() As Variant
    Dim startTime As Double
    Dim endTime As Double
    Dim totalTime As Double

    
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
    
    startTime = Timer 'start time
    'Get the last row of data
    lastRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    Set values = CreateObject("Scripting.Dictionary")
    'Loop through the data in columns A and B
    For i = 1 To lastRow
        'Get the value in column A
        valueA = UCase(ActiveSheet.Cells(i, 1).Value)
        'Get the value in column B
        valueB = ActiveSheet.Cells(i, 2).Value
        'Check if the value in column B already exists in the dictionary
        If Not values.Exists(valueB) Then
            'If it doesn't, add the value in column B as a key and value in column A as the first value in the array
            values.Add valueB, Array(valueA)
        Else
            'If it does, add value in column A into the existing array
            currentArray = values(valueB)
            ReDim Preserve currentArray(UBound(currentArray) + 1)
            currentArray(UBound(currentArray)) = valueA
            values(valueB) = currentArray
        End If
    Next i
    'Loop through the dictionary and save the values in column A as separate CSV files
    For Each Key In values.keys
        'Create a new CSV file
        Set NewFile = CreateObject("Scripting.FileSystemObject").CreateTextFile(myFolder & fileName & "_" & CStr(Key) & ".csv")
        'Write the values in column A to the file
        NewFile.Write Join(values(Key), vbNewLine)
        'Close the file
        NewFile.Close
    Next Key
    endTime = Timer 'end time
    totalTime = endTime - startTime
    Call Shell("explorer.exe" & " " & myFolder, vbNormalFocus)
    MsgBox "Splitter Finished " & myFolder & vbNewLine & " Total time: " & totalTime & " seconds"

    
End Sub
