VERSION 5.00
Begin VB.Form fPen 
   AutoRedraw      =   -1  'True
   Caption         =   "Pen"
   ClientHeight    =   3480
   ClientLeft      =   4635
   ClientTop       =   3300
   ClientWidth     =   5925
   LinkTopic       =   "Form2"
   ScaleHeight     =   3480
   ScaleWidth      =   5925
   Begin VB.CommandButton Command1 
      Caption         =   "Create 1000 random pens"
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   1440
      Width           =   2655
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "fPen.frx":0000
      Left            =   240
      List            =   "fPen.frx":0019
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   2160
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   480
      Width           =   615
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   300
      Left            =   2760
      Max             =   10
      Min             =   1
      TabIndex        =   0
      Top             =   480
      Value           =   1
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "Click and drag around this form to draw with the pen."
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "Width:"
      Height          =   255
      Index           =   1
      Left            =   2160
      TabIndex        =   4
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Pen Type:"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "fPen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mhPen As Long

Private miX As Long
Private miY As Long

Private Sub Combo1_Click()
    pCreatePen
End Sub

Private Sub Command1_Click()
    Dim lhPen(0 To 999) As Long
    Dim i As Long
    Randomize
    For i = 0 To 999
        lhPen(i) = CreatePen(Rnd * (gdiPSInsideFrame + OneL), Rnd * 10, vbBlack)
    Next
    Form1.UpdateStats
'    Do While xKeyIsDown(VK_CONTROL)
'    DoEvents: DoEvents
'    Loop
    For i = 0 To 999
        DeleteObject lhPen(i)
    Next
End Sub

Private Sub Form_Load()
    ScaleMode = vbPixels
    VScroll1.Value = VScroll1.Max
    Combo1.ListIndex = 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Static iX As Long
    Static iY As Long
    Static bInit As Boolean
    
    If Button = 1 Then
        If bInit Then
            Dim lhDc As Long
            Dim lhPenOld As Long
            Dim tPJunk As POINTAPI
            
            lhDc = hdc
            
            If mhPen Then
                lhPenOld = SelectObject(lhDc, mhPen)
                
                MoveToEx lhDc, iX, iY, tPJunk
                LineTo lhDc, x, y
                
                SelectObject lhDc, lhPenOld
            End If
        End If
        Refresh
        Label1(2).Visible = False
        bInit = True
    Else
        bInit = False
    End If
    iX = x
    iY = y
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mhPen Then DeleteObject mhPen
End Sub

Private Sub Text1_Change()
    pCreatePen
End Sub

Private Sub VScroll1_Change()
    Text1.Text = VScroll1.Max - VScroll1.Value + VScroll1.Min
End Sub

Private Sub pCreatePen()
    If mhPen Then DeleteObject mhPen
    mhPen = CreatePen(Combo1.ListIndex, Val(Text1.Text), vbBlack)
End Sub
