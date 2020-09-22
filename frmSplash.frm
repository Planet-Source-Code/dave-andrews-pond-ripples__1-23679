VERSION 5.00
Begin VB.Form frmSplash 
   AutoRedraw      =   -1  'True
   Caption         =   "Splash"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3165
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   3165
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   2160
      Top             =   1680
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      DrawWidth       =   3
      Height          =   2460
      Left            =   0
      ScaleHeight     =   160
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   158
      TabIndex        =   0
      Top             =   0
      Width           =   2430
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Form_Load()
Me.Show
DoEvents
Picture1.Picture = LoadPicture("time.jpg")
InitRipples Picture1.hDC, "time.jpg"
End Sub


Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
AddRipple CLng(x), CLng(y)
Timer1.Interval = 1

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button And 1 Then
    AddRipple CLng(x), CLng(y)
End If
End Sub

Private Sub Timer1_Timer()
RenderRipples
End Sub


