VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


                    '**********************************************
                    '***              quanTest                  ***
                    '***    Programmer : AYON BANERJEE          ***
                    '***        Mobile : 098303 55**3           ***
                    '***    Email : y.ayon@yahoo.co.in          ***
                    '***  Created : 01/08/2006 AT 01:28 PM      ***
                    '**********************************************


Option Explicit

Function pBar(frm As Variant, ByVal progressBarMaxValue As Integer, ByVal loopValue As Integer)

'Show the ProgressBar

Dim X, Y As Integer
frm.ProgressBar1.Max = progressBarMaxValue

For X = 0 To loopValue Step 0.5
    frm.ProgressBar1.Value = X
    For Y = 0 To 500
        DoEvents
    Next Y
    
Next X

If X >= progressBarMaxValue Then
    X = 0
    frm.ProgressBar1.Value = X
End If

End Function

