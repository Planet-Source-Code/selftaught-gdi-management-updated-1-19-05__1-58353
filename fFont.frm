VERSION 5.00
Begin VB.Form fFont 
   Caption         =   "Font"
   ClientHeight    =   3135
   ClientLeft      =   6225
   ClientTop       =   2730
   ClientWidth     =   3360
   HasDC           =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   3135
   ScaleWidth      =   3360
   Begin VB.CommandButton cmd 
      Caption         =   "Create 1000 random fonts"
      Height          =   495
      Index           =   1
      Left            =   600
      TabIndex        =   1
      Top             =   2160
      Width           =   2295
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Browse..."
      Height          =   495
      Index           =   0
      Left            =   1080
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "fFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mhFont As Long
Private WithEvents moFont As cFont
Attribute moFont.VB_VarHelpID = -1

Private Sub cmd_Click(Index As Integer)
    If Index = 0 Then
        If moFont.Browse() Then Refresh
    ElseIf Index = 1 Then
        Dim loFont As cFont
        Set loFont = New cFont
        Dim lhFonts(0 To 999) As Long
        With loFont
            .PutFontInfo Me.Font
            Dim i As Long
            Randomize
            For i = 0 To 999
                .FaceName = Screen.Fonts(Rnd * Screen.FontCount)
                lhFonts(i) = .GetHandle()
            Next
            Form1.UpdateStats
'            Do While xKeyIsDown(VK_CONTROL)
'            DoEvents: DoEvents
'            Loop
            For i = 0 To 999
                .ReleaseHandle lhFonts(i)
            Next
        End With
    End If
End Sub

Private Sub Form_Load()
    Set moFont = New cFont
    moFont.PutFontInfo Me.Font
End Sub

Private Sub Form_Paint()
    Dim lhFont As Long
    Dim lhDc As Long
    lhDc = hdc
    
    Cls
    
    If mhFont Then
        lhFont = SelectObject(lhDc, mhFont)
        TextOut lhDc, 20, 20, "This is a test string!", 22&
        SelectObject lhDc, lhFont
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mhFont Then moFont.ReleaseHandle mhFont
    Set moFont = Nothing
End Sub

Private Sub moFont_Changed()
    If mhFont Then moFont.ReleaseHandle mhFont
    mhFont = moFont.GetHandle()
End Sub
