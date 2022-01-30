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
            
            'Loop through 2nd column as long as there is some data (consider merged cells)
            Do While GetMergedValue(selectedWorksheet.Cells(i, 2)) <> ""
                MsgBox GetMergedValue(selectedWorksheet.Cells(i, 2))
                i = i + 1
            Loop

        End If

    Next chkBox

End Sub
