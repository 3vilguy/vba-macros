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
            
            'Create row counter for selected worksheet
            Dim i As Integer
            i = 1
            
            'Loop through 1st column as long as there is some data
            Do While selectedWorksheet.Cells(i, 1).Value <> ""
                MsgBox selectedWorksheet.Cells(i, 1).Value
                i = i + 1
            Loop

        End If

    Next chkBox

End Sub
