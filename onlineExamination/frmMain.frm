VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Welcome to 'O' level online examination test."
   ClientHeight    =   7335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   200
      Left            =   14640
      Top             =   480
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   14280
      Top             =   8160
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   180
      Left            =   120
      TabIndex        =   33
      Top             =   7080
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   318
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Frame Frame2 
      Height          =   6975
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   15015
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Height          =   3015
         Left            =   9600
         TabIndex        =   34
         Top             =   3480
         Width           =   5175
         Begin VB.Shape Shape3 
            Height          =   2655
            Left            =   120
            Top             =   240
            Width           =   4935
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Decision Box"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   37
            Top             =   360
            Width           =   4455
         End
         Begin VB.Line Line5 
            X1              =   120
            X2              =   5040
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Line Line6 
            X1              =   120
            X2              =   5040
            Y1              =   2400
            Y2              =   2400
         End
         Begin VB.Label lblremain 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   2565
            TabIndex        =   36
            Top             =   2520
            Width           =   45
         End
         Begin VB.Label lblAnswer 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2460
            TabIndex        =   35
            Top             =   1320
            Width           =   135
         End
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   4575
         Left            =   240
         TabIndex        =   24
         Top             =   2040
         Width           =   9255
         Begin VB.OptionButton optAnswer 
            Height          =   255
            Index           =   0
            Left            =   360
            MouseIcon       =   "frmMain.frx":56CA
            MousePointer    =   99  'Custom
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   1800
            Width           =   8655
         End
         Begin VB.OptionButton optAnswer 
            Height          =   255
            Index           =   1
            Left            =   360
            MouseIcon       =   "frmMain.frx":59D4
            MousePointer    =   99  'Custom
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   2520
            Width           =   8655
         End
         Begin VB.OptionButton optAnswer 
            Height          =   255
            Index           =   2
            Left            =   360
            MouseIcon       =   "frmMain.frx":5CDE
            MousePointer    =   99  'Custom
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   3240
            Width           =   8655
         End
         Begin VB.OptionButton optAnswer 
            Height          =   255
            Index           =   3
            Left            =   360
            MouseIcon       =   "frmMain.frx":5FE8
            MousePointer    =   99  'Custom
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   3960
            Width           =   8655
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "d."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   41
            Top             =   3990
            Width           =   180
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "c."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   40
            Top             =   3270
            Width           =   180
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "b."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   39
            Top             =   2550
            Width           =   180
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "a."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   1800
            Width           =   195
         End
         Begin VB.Shape Shape2 
            Height          =   735
            Left            =   0
            Top             =   480
            Width           =   9135
         End
         Begin VB.Label Label12 
            Caption         =   "Questions :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label lblQtn 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   120
            TabIndex        =   31
            Top             =   600
            Width           =   8835
         End
         Begin VB.Label lblQNo 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1560
            TabIndex        =   30
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label Label13 
            Caption         =   "Answer :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Shape shp 
            Height          =   495
            Index           =   0
            Left            =   0
            Top             =   1680
            Width           =   9135
         End
         Begin VB.Shape shp 
            Height          =   495
            Index           =   1
            Left            =   0
            Top             =   2400
            Width           =   9135
         End
         Begin VB.Shape shp 
            Height          =   495
            Index           =   2
            Left            =   0
            Top             =   3120
            Width           =   9135
         End
         Begin VB.Shape shp 
            Height          =   495
            Index           =   3
            Left            =   0
            Top             =   3840
            Width           =   9135
         End
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   14580
         TabIndex        =   23
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmdStart 
         Caption         =   "&Start"
         Height          =   375
         Left            =   4920
         TabIndex        =   22
         Top             =   1320
         Width           =   615
      End
      Begin VB.Shape Shape1 
         Height          =   615
         Left            =   4800
         Top             =   1200
         Width           =   855
      End
      Begin VB.Line Line4 
         X1              =   240
         X2              =   5760
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label lblGetNo 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   4320
         TabIndex        =   21
         Top             =   1560
         Width           =   90
      End
      Begin VB.Label lblTotalNo 
         AutoSize        =   -1  'True
         Caption         =   "num"
         Height          =   195
         Left            =   4320
         TabIndex        =   20
         Top             =   1200
         Width           =   300
      End
      Begin VB.Label lblSubject 
         AutoSize        =   -1  'True
         Caption         =   "sub"
         Height          =   195
         Left            =   1320
         TabIndex        =   19
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label Label11 
         Caption         =   "Score              :"
         Height          =   255
         Left            =   3120
         TabIndex        =   18
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Total Number  :"
         Height          =   255
         Left            =   3120
         TabIndex        =   17
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Time        :    5 minutes."
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label8 
         Caption         =   "Subject    :"
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label lblStudentName 
         BackStyle       =   0  'Transparent
         Caption         =   "Ayon Banerjee"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1920
         TabIndex        =   14
         Top             =   600
         Width           =   4680
      End
      Begin VB.Line Line3 
         X1              =   240
         X2              =   5760
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome,"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   13
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "'O' LEVEL EXAMINATION"
         BeginProperty Font 
            Name            =   "Vineta BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   5655
      End
      Begin VB.Line Line2 
         X1              =   240
         X2              =   5760
         Y1              =   600
         Y2              =   600
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15015
      Begin VB.CommandButton cmdSubmit 
         Caption         =   "&Submit"
         Enabled         =   0   'False
         Height          =   855
         Left            =   5160
         TabIndex        =   10
         Top             =   2280
         Width           =   1095
      End
      Begin VB.ComboBox cmbSubject 
         BackColor       =   &H8000000F&
         Height          =   315
         ItemData        =   "frmMain.frx":62F2
         Left            =   3120
         List            =   "frmMain.frx":6305
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2760
         Width           =   1935
      End
      Begin VB.TextBox txtPhone 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   3120
         MaxLength       =   250
         TabIndex        =   8
         ToolTipText     =   "Your Phone Number"
         Top             =   2280
         Width           =   1935
      End
      Begin VB.TextBox txtAddress 
         BackColor       =   &H8000000F&
         Height          =   765
         Left            =   3120
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Tag             =   "Required"
         ToolTipText     =   "Your Address"
         Top             =   1320
         Width           =   5655
      End
      Begin VB.TextBox txtName 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   3120
         MaxLength       =   250
         TabIndex        =   6
         Tag             =   "Required"
         ToolTipText     =   "Your Name"
         Top             =   840
         Width           =   5655
      End
      Begin VB.Label Label5 
         Caption         =   "Select Subject                             :"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   2775
         Width           =   2415
      End
      Begin VB.Label Label4 
         Caption         =   "Enter your Phone No.                  :"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   2295
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "Enter your Address                       :"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Enter your Name                          :"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   855
         Width           =   2415
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   5760
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label1 
         Caption         =   "'O' LEVEL EXAMINATION"
         BeginProperty Font 
            Name            =   "Vineta BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   5655
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

                    '**********************************************
                    '***              quanTest                  ***
                    '***    Programmer : AYON BANERJEE          ***
                    '***        Mobile : 098303 55**3           ***
                    '***    Email : y.ayon@yahoo.co.in          ***
                    '***  Created : 01/08/2006 AT 01:28 PM      ***
                    '**********************************************



Option Explicit
Dim studName As String, studAdd As String, studPhone As String
Dim subject As String
Dim tableName As String
Dim rightAnswer As Integer
Dim marks As Integer, noofQ As Integer
Dim flagStat As Boolean
Dim objSelectQuery As New ClSelectQuery
Dim startTime As Integer, endTime As Integer

Dim strCaption As String
Dim i, v As Integer

Private Sub cmbSubject_Click()
    
    If cmbSubject.ListIndex >= 0 Then
        cmdSubmit.Enabled = True
    End If
    
    If cmbSubject.ListIndex = 0 Then
        tableName = "qaIT"
    ElseIf cmbSubject.ListIndex = 1 Then
        tableName = "qaDOS"
    ElseIf cmbSubject.ListIndex = 2 Then
        tableName = "qaWindows"
    ElseIf cmbSubject.ListIndex = 3 Then
        tableName = "qaWord"
    ElseIf cmbSubject.ListIndex = 4 Then
        tableName = "qaSpreadsheet"
    End If

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdStart_Click()
    cmdStart.Enabled = False
    
    
    objSelectQuery.SelectQuery "select * from " + tableName + ""
    
    If Not objSelectQuery.rs.EOF = True Then
        lblQNo.Caption = objSelectQuery.rs("slNo")
        lblQtn.Caption = objSelectQuery.rs("questions")
        optAnswer(0).Caption = objSelectQuery.rs("answer1")
        optAnswer(1).Caption = objSelectQuery.rs("answer2")
        optAnswer(2).Caption = objSelectQuery.rs("answer3")
        optAnswer(3).Caption = objSelectQuery.rs("answer4")
        rightAnswer = objSelectQuery.rs("rightAnswer")
    End If
    
    Frame3.Enabled = True
    startTime = Minute(Time)
    endTime = startTime + 5
    Timer1.Enabled = True

End Sub

Private Sub cmdSubmit_Click()
    ' check for required fileds
    Call requiredFields(Me)
        If setValue = "True" Then
            Exit Sub
        End If

studName = Trim(txtName.Text)
studAdd = Trim(txtAddress.Text)
studPhone = Trim(txtPhone.Text)
subject = Trim(cmbSubject.Text)


Frame1.Visible = False
Frame2.Visible = True

lblStudentName.Caption = studName
lblSubject.Caption = subject

    Dim objSelectQuery2 As New ClSelectQuery
    objSelectQuery2.SelectQuery "select * from " + tableName + ""
    
    While Not objSelectQuery2.rs.EOF = True
        noofQ = noofQ + 1
        objSelectQuery2.rs.MoveNext
    Wend
    
    objSelectQuery2.closeSelectQuery

    lblTotalNo.Caption = noofQ * 10

End Sub

Private Sub Form_Load()
marks = 0
noofQ = 0
flagStat = True

strCaption = "Welcome to 'O' level online examination test.        "
v = Len(strCaption)

End Sub

Private Sub optAnswer_Click(Index As Integer)

If flagStat = True Then
    Dim answer As Integer
    Dim progressBar As New ClProgressBar
    
    Me.MousePointer = vbHourglass
    Frame3.Enabled = False

    progressBar.pBar Me, 100, 100    'show the progressBar

    answer = Val(Index) + 1
        
    If rightAnswer = answer Then
        lblAnswer.Caption = "RIGHT ANSWER"
        lblAnswer.ForeColor = RGB(0, 255, 0)
        marks = marks + 10
        lblGetNo.Caption = marks
    Else
        lblAnswer.Caption = "WRONG ANSWER"
        lblAnswer.ForeColor = RGB(255, 0, 0)
    End If


    objSelectQuery.rs.MoveNext
    
    If Not objSelectQuery.rs.EOF = True Then
        lblQNo.Caption = objSelectQuery.rs("slNo")
        lblQtn.Caption = objSelectQuery.rs("questions")
        optAnswer(0).Caption = objSelectQuery.rs("answer1")
        optAnswer(1).Caption = objSelectQuery.rs("answer2")
        optAnswer(2).Caption = objSelectQuery.rs("answer3")
        optAnswer(3).Caption = objSelectQuery.rs("answer4")
        rightAnswer = objSelectQuery.rs("rightAnswer")
    End If
    
    optAnswer(Index).Value = False
    
    progressBar.pBar Me, 100, 100    'show the progressBar
    lblAnswer.Caption = ""
    
    Me.MousePointer = vbNormal
    Frame3.Enabled = True

    If objSelectQuery.rs.EOF = True Then
        objSelectQuery.closeSelectQuery
        flagStat = False
        Frame3.Enabled = False
        Call showResult
    End If
    
End If
    
End Sub

Private Sub Timer1_Timer()
    lblremain.Caption = "Time remaining  :  " & Str(endTime - Minute(Time)) & " minute"
    If (endTime - Minute(Time)) <= 0 Then
                MsgBox "Sorry !! You have no more time to answer these questions.", vbInformation + vbOKOnly
                objSelectQuery.closeSelectQuery
                flagStat = False
                Frame3.Enabled = False
                Call showResult
    End If

End Sub

Function showResult()
Frame3.Visible = False
Timer1.Enabled = False

Dim leftPos As Integer, topPos As Integer
Dim delayLoop As Integer

For leftPos = 9600 To 120 Step -1

    Frame4.Left = leftPos
    
    For delayLoop = 1 To 25
        DoEvents
    Next delayLoop
    
Next leftPos

For topPos = 3480 To 3000 Step -1

    Frame4.Top = topPos
    
    For delayLoop = 1 To 25
        DoEvents
    Next delayLoop
    
Next topPos

lblAnswer.FontSize = 12
lblAnswer.ForeColor = RGB(0, 0, 0)

If Val(lblGetNo.Caption) >= (Val(lblTotalNo.Caption) * 60) / 100 Then
    lblAnswer.Caption = "Congratulations !! " & studName & vbNewLine & "You are qualified to 'A' level."
    lblremain.Caption = "You have get " & lblGetNo.Caption & " out of " & lblTotalNo.Caption
Else
    lblAnswer.Caption = "Sorry !! " & studName & vbNewLine & "You are not qualified to 'A' level."
    lblremain.Caption = "You have get " & lblGetNo.Caption & " out of " & lblTotalNo.Caption

End If


End Function

Private Sub Timer2_Timer()
    Me.Caption = Left$(strCaption, i)
    i = i + 1
    If i = v Then
        i = 0
    End If

End Sub
