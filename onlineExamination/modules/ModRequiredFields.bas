Attribute VB_Name = "ModRequiredFields"

                    '**********************************************
                    '***              quanTest                  ***
                    '***    Programmer : AYON BANERJEE          ***
                    '***        Mobile : 098303 55**3           ***
                    '***    Email : y.ayon@yahoo.co.in          ***
                    '***  Created : 01/08/2006 AT 01:28 PM      ***
                    '**********************************************




Option Explicit
Public regConnectionString As String
Public setValue As String
Public colourCode As String


Public Function requiredFields(frm As Variant)
    Dim Y As Integer
    For Y = 0 To frm.Controls.Count - 1
        If TypeOf frm.Controls(Y) Is TextBox Then
            If frm.Controls(Y).Text = "" And frm.Controls(Y).Tag = "Required" Then
                colourCode = Str(frm.Controls(Y).BackColor)
                frm.Controls(Y).BackColor = &HFF&
                MsgBox frm.Controls(Y).ToolTipText + " must be required.", vbCritical + vbOKOnly
                frm.Controls(Y).SetFocus
                frm.Controls(Y).BackColor = colourCode
                setValue = "True"
                Exit Function
            Else
                setValue = "False"
            End If
        ElseIf TypeOf frm.Controls(Y) Is ComboBox Then
            If frm.Controls(Y).Text = "" And frm.Controls(Y).Tag = "Required" Then
                colourCode = Str(frm.Controls(Y).BackColor)
                frm.Controls(Y).BackColor = &HFF&
                MsgBox frm.Controls(Y).ToolTipText + " must be required.", vbCritical + vbOKOnly
                frm.Controls(Y).SetFocus
                frm.Controls(Y).BackColor = colourCode
                setValue = "True"
                Exit Function
            Else
                setValue = "False"
            End If
       ' ElseIf TypeOf frm.Controls(y) Is DataGrid Then
            'write code to reset DataGrid
            'MsgBox "Remember to call preResetAll first", vbOKOnly, "TO DO"
          '  frm.Controls(y).ReBind
          '  frm.Controls(y).Refresh
      ' ElseIf TypeOf frm.Controls(y) Is DTPicker Then
        '    frm.Controls(y).value = Date
       ElseIf TypeOf frm.Controls(Y) Is ListBox Then
            If frm.Controls(Y).ListIndex = -1 And frm.Controls(Y).Tag = "Required" Then
                colourCode = Str(frm.Controls(Y).BackColor)
                frm.Controls(Y).BackColor = &HFF&
                MsgBox frm.Controls(Y).ToolTipText + " must be required.", vbCritical + vbOKOnly
                frm.Controls(Y).SetFocus
                frm.Controls(Y).BackColor = colourCode
                setValue = "True"
                Exit Function
            Else
                setValue = "False"
            End If
        End If
    Next Y

End Function
