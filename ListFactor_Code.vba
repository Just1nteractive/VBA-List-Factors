Private Sub CommandButton1_Click()
Unload Me   'Close button
End Sub

Private Sub CommandButton2_Click()
If IsNumeric(TextBox1) Then 'Check if the value is numeric...
    Dim num                 'This is how we will reference the number
    num = Abs(Int(TextBox1.Value))  'Get the number, in number form (instead of string) and make Absolutely sure it's positive
    TextBox1.Value = num            'Replace the textbox value with the converted, so we know what we're dealing with
    If num > 2 Then     'Check if the number is greater than 2
        ListBox1.Clear  'Clear any previous lists, so as not to keep adding to the list.
        DoEvents        'This allows the controls time to fresh
        'num = Int(TextBox1.Value)   'Get the number, in number form (instead of string)
        For x = 2 To (num * 0.5)    'Define the loop, we only need to verify until the halfway point since nothing above 5 will be a factor for 10.
            If num Mod x = 0 Then   'If it divides evenly, no remainder.. that means, it's a factor!
                ListBox1.AddItem x  'Add the iteration (x) to the list
                DoEvents               'Refresh the list
            End If
            ProgressBar1.Value = (x / (num * 0.5)) * 100  'Update the progress bar
        Next                        'Keep moving
        ProgressBar1.Value = 100
    Else                            'If the number is less than two, advise the user:
        MsgBox "This number is too low...", vbOKOnly, "Give me a challenge"
    End If
Else                                'If something other than a number is entered
    MsgBox "That's not a recognized NUMBER", vbOKOnly, "Numbers Only"
End If
End Sub

Private Sub CommandButton3_Click()
ListBox1.Clear  'Reset button
TextBox1.Text = ""
ProgressBar1.Value = 0
End Sub
