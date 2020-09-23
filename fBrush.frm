VERSION 5.00
Begin VB.Form fBrush 
   Caption         =   "Form2"
   ClientHeight    =   4755
   ClientLeft      =   4155
   ClientTop       =   3240
   ClientWidth     =   5985
   LinkTopic       =   "Form2"
   ScaleHeight     =   4755
   ScaleWidth      =   5985
   Begin VB.CommandButton Command1 
      Caption         =   "Create 1000 random brushes"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   2655
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      HasDC           =   0   'False
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   5925
      TabIndex        =   0
      Top             =   0
      Width           =   5985
   End
   Begin VB.Label lbl 
      Caption         =   "Click on a color to create a brush..."
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   3255
   End
End
Attribute VB_Name = "fBrush"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long

Private mhBrush As Long

Private Sub Command1_Click()
    Dim lhBrush(0 To 999) As Long
    Dim i As Long
    Randomize
    For i = 0 To 999
        lhBrush(i) = CreateSolidBrush(RGB(45 + Rnd * 20, 120 + 10 * Rnd, 170 + 5 * Rnd))
    Next
    Form1.UpdateStats
'    Do While xKeyIsDown(VK_CONTROL)
'    DoEvents: DoEvents
'    Loop
    For i = 0 To 999
        DeleteObject lhBrush(i)
    Next
End Sub

Private Sub Form_Load()
    ScaleMode = vbPixels
    Picture1.ScaleMode = vbPixels
End Sub

Private Sub Form_Paint()
    Cls
    If mhBrush Then
        Dim tR As RECT
        tR.Bottom = ScaleHeight
        tR.Right = ScaleWidth
        FillRect hdc, tR, mhBrush
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mhBrush Then DeleteObject mhBrush
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If mhBrush Then DeleteObject mhBrush
        mhBrush = CreateSolidBrush(Picture1.Point(x, y))
        lbl.Visible = False
        Refresh
        Picture1.Refresh
    End If
End Sub

Private Sub Picture1_Paint()
    Dim i As Long
    Dim tR As RECT
    Dim liWidth As Long
    liWidth = Picture1.Width
    For i = ZeroL To 15
        With tR
            .Left = (liWidth * i) \ 16
            .Top = ZeroL
            .Right = (liWidth * (i + OneL)) \ 16
            .Bottom = Picture1.ScaleHeight
            Picture1.Line (.Left, .Top)-(.Right, .Bottom), QBColor(i), BF
        End With
    Next
End Sub

Private Sub Picture1_Resize()
    Picture1.Refresh
End Sub
