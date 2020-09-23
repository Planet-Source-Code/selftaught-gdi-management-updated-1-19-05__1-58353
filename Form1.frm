VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2820
   ClientLeft      =   3690
   ClientTop       =   2850
   ClientWidth     =   5865
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2820
   ScaleWidth      =   5865
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   1680
      Top             =   1320
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Pen"
      Height          =   495
      Index           =   2
      Left            =   480
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Brush"
      Height          =   495
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Font"
      Height          =   495
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Caption         =   "Created:"
      Height          =   255
      Index           =   10
      Left            =   4560
      TabIndex        =   13
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Caption         =   "Requested:"
      Height          =   255
      Index           =   9
      Left            =   3360
      TabIndex        =   12
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   8
      Left            =   4560
      TabIndex        =   11
      Top             =   600
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   7
      Left            =   4560
      TabIndex        =   10
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   6
      Left            =   4560
      TabIndex        =   9
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   5
      Left            =   3360
      TabIndex        =   8
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   4
      Left            =   3360
      TabIndex        =   7
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   3
      Left            =   3360
      TabIndex        =   6
      Top             =   600
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Pens:"
      Height          =   255
      Index           =   2
      Left            =   2040
      TabIndex        =   5
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Brushes:"
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   4
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Fonts:"
      Height          =   255
      Index           =   0
      Left            =   2040
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub cmd_Click(Index As Integer)
    Dim loForm As Form
    Select Case Index
        Case 0: Set loForm = New fFont
        Case 1: Set loForm = New fBrush
        Case 2: Set loForm = New fPen
    End Select
    loForm.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim F As Form
    For Each F In Forms
        If Not F Is Me Then Unload F
    Next
End Sub

Private Sub Timer1_Timer()
    Dim li(0 To 5) As Long
    mGDI.Statistics li(0), li(5), li(1), li(4), li(2), li(3)
    Dim i As Long
    For i = 0 To 5
        lbl(i + 3).Caption = li(i)
    Next
End Sub

Public Sub UpdateStats()
    Timer1_Timer
End Sub
