Sub Splitter()
    
    
    
    Dim FldrPicker As FileDialog
    Dim myFolder As String
    Dim fileName As String

    
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
    
    startTime = Now()
    'Get the last row of data
    lastRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row

    'Create an empty dictionary to store values
    Dim values As Object
    Set values = CreateObject("Scripting.Dictionary")

    'Loop through the data in columns A and B
    For i = 1 To lastRow
        'Get the value in column A
        valueA = UCase(ActiveSheet.Cells(i, 1).Value)
    
        'Get the value in column B
        valueB = ActiveSheet.Cells(i, 2).Value
    
        'Check if the value in column A already exists in the dictionary
        If Not values.Exists(valueA) Then
            'If it doesn't, add the value in column A as a key and value in column B as the first value
            values.Add valueA, valueB
            'If it does, skip the line
            
        End If
    Next i
    
    'Removes 0 from split
    For Each Key In values.keys
        If values(Key) = 0 Then
            values.Remove Key
        End If
    Next Key


    'Removes header row if it exists
    If values.Exists(UCase("Acct")) Then
        values.Remove (UCase("Acct"))
    End If
        
    
    For Each Key In values.keys
        If CreateObject("Scripting.FileSystemObject").FileExists(myFolder & fileName & "_" & CStr(values(Key)) & ".csv") Then
            Set NewFile = CreateObject("Scripting.FileSystemObject").OpenTextFile(myFolder & fileName & "_" & CStr(values(Key)) & ".csv", 8)
            'Write the values in column A to the file
            NewFile.WriteLine (Key)
        Else
            'Create a new CSV file
            Set NewFile = CreateObject("Scripting.FileSystemObject").CreateTextFile(myFolder & fileName & "_" & CStr(values(Key)) & ".csv")
            'Write the values in column A to the file
            NewFile.WriteLine (Key)
        End If
        'Close the file
        NewFile.Close
    Next Key
    
    endTime = Now()
    totalTime = DateDiff("s", startTime, endTime)
    Call Shell("explorer.exe" & " " & myFolder, vbNormalFocus)
    MsgBox "Splitter Finished " & myFolder & vbNewLine & " Total time: " & totalTime & " seconds"
    
End Sub
