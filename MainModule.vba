Sub Button1_Click()

    'Create variable for checkbox
    Dim chkBox As CheckBox

    'Loop through each checkbox on active sheet
    For Each chkBox In ActiveSheet.CheckBoxes

        'Check if checkbox is selected
        If chkBox.Value = 1 Then

            'Set worksheet reference (based on selected checkbox)
            Dim selectedWorksheet As Worksheet
            Set selectedWorksheet = Worksheets(chkBox.Caption)
            
            'Create row counter for selected worksheet (start from row number 3)
            Dim i As Integer
            i = 3
            
            'Create other variables used in the loop
            Dim deliverableTitle As String
            
            'Loop through 2nd column as long as there is some data (consider merged cells)
            Do While Not IsEmpty(GetMergedValue(selectedWorksheet.Cells(i, 2)))

                'Check if it's deliverable title (1st column would be empty)
                If IsEmpty(GetMergedValue(selectedWorksheet.Cells(i, 1))) Then
                    'It is!
                    deliverableTitle = GetMergedValue(selectedWorksheet.Cells(i, 2))
                Else
                    'It's not. Let's check the commodity column
                    If Not IsEmpty(selectedWorksheet.Cells(i, 3)) Then
                    
                        'Something was selected from the dropdown
                        MsgBox GetMergedValue(selectedWorksheet.Cells(i, 1)) & ", " & deliverableTitle & ", " & selectedWorksheet.Cells(i, 3)
                    End If
                End If

                i = i + 1
            Loop

        End If

    Next chkBox

End Sub
