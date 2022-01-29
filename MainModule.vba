Sub Button1_Click()

    'Create variable for checkbox
    Dim chkBox As CheckBox

    'Loop through each checkbox on active sheet
    For Each chkBox In ActiveSheet.CheckBoxes

        'Check if checkbox is selected
        If chkBox.Value = 1 Then
            MsgBox chkBox.Caption
        End If

    Next chkBox

End Sub
