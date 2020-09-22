VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000A&
   Caption         =   "Electronic Etch-A-Sketch"
   ClientHeight    =   6345
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6990
   ForeColor       =   &H8000000C&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6345
   ScaleWidth      =   6990
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Speed"
      Height          =   1455
      Left            =   2640
      TabIndex        =   23
      Top             =   4320
      Width           =   1815
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Change S&peed"
         Default         =   -1  'True
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         Caption         =   "/\"
         Height          =   315
         Left            =   840
         TabIndex        =   25
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton Command4 
         Caption         =   "\/"
         Height          =   315
         Left            =   1200
         TabIndex        =   24
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Current:"
         Height          =   255
         Left            =   480
         TabIndex        =   29
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "0.04"
         Height          =   255
         Left            =   1080
         TabIndex        =   28
         Top             =   1080
         Width           =   615
      End
   End
   Begin VB.CommandButton Command1 
      DownPicture     =   "form1.frx":0000
      Height          =   495
      Index           =   9
      Left            =   6120
      Picture         =   "form1.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4320
      Width           =   615
   End
   Begin VB.OptionButton optcol 
      BackColor       =   &H0000FFFF&
      Caption         =   "Yellow"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   2400
      TabIndex        =   21
      Top             =   5880
      Width           =   855
   End
   Begin VB.OptionButton optcol 
      BackColor       =   &H00000000&
      Caption         =   "Black"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   3240
      TabIndex        =   20
      Top             =   5880
      Width           =   735
   End
   Begin VB.OptionButton optcol 
      BackColor       =   &H000000FF&
      Caption         =   "Red"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   1680
      TabIndex        =   19
      Top             =   5880
      Width           =   735
   End
   Begin VB.OptionButton optcol 
      BackColor       =   &H0000FF00&
      Caption         =   "Green"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   840
      TabIndex        =   18
      Top             =   5880
      Width           =   855
   End
   Begin VB.OptionButton optcol 
      BackColor       =   &H00FF0000&
      Caption         =   "Blue"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   17
      Top             =   5880
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      DownPicture     =   "form1.frx":0884
      Height          =   495
      Index           =   12
      Left            =   4560
      Picture         =   "form1.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4320
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      DownPicture     =   "form1.frx":1108
      Height          =   495
      Index           =   11
      Left            =   4560
      Picture         =   "form1.frx":154A
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5520
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      DownPicture     =   "form1.frx":198C
      Height          =   495
      Index           =   10
      Left            =   6120
      Picture         =   "form1.frx":1DCE
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5520
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Reset"
      Height          =   375
      Index           =   8
      Left            =   1200
      TabIndex        =   12
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Clear Screen"
      Height          =   375
      Index           =   7
      Left            =   1200
      TabIndex        =   11
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      DownPicture     =   "form1.frx":2210
      Height          =   615
      Index           =   6
      Left            =   5280
      Picture         =   "form1.frx":2652
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5400
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      DownPicture     =   "form1.frx":2A94
      Height          =   615
      Index           =   5
      Left            =   5280
      Picture         =   "form1.frx":2ED6
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4320
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      DownPicture     =   "form1.frx":3318
      Height          =   495
      Index           =   4
      Left            =   6000
      Picture         =   "form1.frx":375A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4920
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      DownPicture     =   "form1.frx":3B9C
      Height          =   495
      Index           =   3
      Left            =   4560
      Picture         =   "form1.frx":3FDE
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4920
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Quit"
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   5280
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "S&top"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Start"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "By Daniel Chiti, April 2000. e-mail: dchiti@hotmail.com"
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   6120
      Width           =   6615
   End
   Begin VB.Line Line4 
      X1              =   480
      X2              =   6480
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line8 
      X1              =   0
      X2              =   6960
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line7 
      X1              =   0
      X2              =   0
      Y1              =   4200
      Y2              =   0
   End
   Begin VB.Line Line6 
      X1              =   6960
      X2              =   6960
      Y1              =   4200
      Y2              =   0
   End
   Begin VB.Line Line5 
      X1              =   0
      X2              =   6960
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
      Height          =   2775
      Left            =   600
      TabIndex        =   13
      Top             =   720
      Width           =   5775
   End
   Begin VB.Line Line3 
      X1              =   480
      X2              =   6480
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line2 
      X1              =   6480
      X2              =   6480
      Y1              =   600
      Y2              =   3600
   End
   Begin VB.Line Line1 
      X1              =   480
      X2              =   480
      Y1              =   600
      Y2              =   3600
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   4680
      Shape           =   3  'Circle
      Top             =   3720
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   1080
      Shape           =   3  'Circle
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label3 
      BackColor       =   &H000000C0&
      Height          =   3615
      Left            =   6480
      TabIndex        =   9
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000C0&
      Height          =   3615
      Left            =   0
      TabIndex        =   8
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "Electronic Etch-A-Sketch"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   615
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   6975
   End
   Begin VB.Label Label4 
      BackColor       =   &H000000C0&
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   480
      TabIndex        =   10
      Top             =   3600
      Width           =   6015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************
'Coded By Daniel Chiti              *
'Completed April 18, 2000 at 6:20PM *
'E-Mail: dchiti@hotmail.com         *
'************************************
Dim drawline As Double
Dim drawline2 As Double
Dim endline As Integer
Dim increaser As Double
Dim col As OLE_COLOR
Dim goingback As Double
Dim imgb As Integer
Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
endline = 0
Do Until endline <> 0 Or drawline >= 3550
    PSet (drawline2, drawline), col
    drawline = drawline + increaser
    DoEvents
Loop
Case 1
endline = 1
Case 2
End
Case 3
endline = 3
Do Until endline <> 3 Or drawline2 <= 525
        PSet (drawline2, drawline), col
        drawline2 = drawline2 - increaser
        DoEvents
Loop
Case 4
endline = 4
Do Until endline <> 4 Or drawline2 >= 6400
    PSet (drawline2, drawline), col
    drawline2 = drawline2 + increaser
    goingback = drawline2
    imgb = 4
    DoEvents
Loop
Case 5
endline = 5
Do Until endline <> 5 Or drawline <= 650
    PSet (drawline2, drawline), col
    drawline = drawline - increaser
    DoEvents
Loop
Case 6
endline = 0
Do Until endline <> 0 Or drawline >= 3495
    PSet (drawline2, drawline), col
    drawline = drawline + increaser
    DoEvents
Loop
Case 7
Cls
Case 8
Cls
drawline = 625
drawline2 = 525
increaser = 0.04
Label7.Caption = increaser
Case 9
endline = 9
Do Until endline <> 9 Or drawline <= 650 Or drawline2 >= 6400
    PSet (drawline2, drawline), col
    drawline = drawline - increaser
    drawline2 = drawline2 + increaser
    DoEvents
Loop
Case 10
endline = 10
Do Until endline <> 10 Or drawline >= 3495 Or drawline2 >= 6400
    PSet (drawline2, drawline), col
    drawline = drawline + increaser
    drawline2 = drawline2 + increaser
    DoEvents
Loop
Case 11
endline = 11
Do Until endline <> 11 Or drawline >= 3495 Or drawline2 <= 525
    PSet (drawline2, drawline), col
    drawline = drawline + increaser
    drawline2 = drawline2 - increaser
    DoEvents
Loop
Case 12
endline = 12
Do Until endline <> 12 Or drawline <= 650 Or drawline2 <= 525
    PSet (drawline2, drawline), col
    drawline = drawline - increaser
    drawline2 = drawline2 - increaser
    DoEvents
Loop
End Select
End Sub

Private Sub Command1_LostFocus(Index As Integer)
Select Case Index
Case 0
Command1(0).BackColor = &H8000000A
Case 1
Command1(1).BackColor = &H8000000A
Case 2
Command1(2).BackColor = &H8000000A
Case 3
Command1(3).BackColor = &H8000000A
Case 4
Command1(4).BackColor = &H8000000A
Case 5
Command1(5).BackColor = &H8000000A
Case 6
Command1(6).BackColor = &H8000000A
Case 7
Command1(7).BackColor = &H8000000A
Case 8
Command1(8).BackColor = &H8000000A
Case 9
Command1(9).BackColor = &H8000000A
Case 10
Command1(10).BackColor = &H8000000A
Case 11
Command1(11).BackColor = &H8000000A
Case 12
Command1(12).BackColor = &H8000000A
End Select
End Sub

Private Sub Command1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
Case 0
Command1(0).BackColor = vbWhite
Case 1
Command1(1).BackColor = vbWhite
Case 2
Command1(2).BackColor = vbWhite
Case 3
Command1(3).BackColor = vbWhite
Case 4
Command1(4).BackColor = vbWhite
Case 5
Command1(5).BackColor = vbWhite
Case 6
Command1(6).BackColor = vbWhite
Case 7
Command1(7).BackColor = vbWhite
Case 8
Command1(8).BackColor = vbWhite
Case 9
Command1(9).BackColor = vbWhite
Case 10
Command1(10).BackColor = vbWhite
Case 11
Command1(11).BackColor = vbWhite
Case 12
Command1(12).BackColor = vbWhite
End Select
End Sub

Private Sub Command1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
Case 0
Command1(0).BackColor = &H8000000A
Case 1
Command1(1).BackColor = &H8000000A
Case 2
Command1(2).BackColor = &H8000000A
Case 3
Command1(3).BackColor = &H8000000A
Case 4
Command1(4).BackColor = &H8000000A
Case 5
Command1(5).BackColor = &H8000000A
Case 6
Command1(6).BackColor = &H8000000A
Case 7
Command1(7).BackColor = &H8000000A
Case 8
Command1(8).BackColor = &H8000000A
Case 9
Command1(9).BackColor = &H8000000A
Case 10
Command1(10).BackColor = &H8000000A
Case 11
Command1(11).BackColor = &H8000000A
Case 12
Command1(12).BackColor = &H8000000A
End Select
End Sub

Private Sub Command2_Click()
increaser = Text1.Text
Text1.Text = ""
Label7.Caption = increaser
End Sub

Private Sub Data1_Validate(Action As Integer, Save As Integer)

End Sub


Private Sub Command3_Click()
increaser = increaser + 0.01
Label7.Caption = increaser
Beep
End Sub

Private Sub Command4_Click()
increaser = increaser - 0.01
Label7.Caption = increaser
Beep
End Sub

Private Sub Form_Load()
drawline = 625
drawline2 = 525
increaser = 0.04
End Sub

Private Sub optcol_Click(Index As Integer)
Select Case Index
Case 0
col = vbBlue
Case 1
col = vbGreen
Case 2
col = vbRed
Case 3
col = vbYellow
Case 4
col = vbBlack
End Select
End Sub
