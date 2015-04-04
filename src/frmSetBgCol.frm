VERSION 5.00
Begin VB.Form frmSetBgCol 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Set BG Color"
   ClientHeight    =   1185
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3045
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   2  'Cross
   ScaleHeight     =   1185
   ScaleWidth      =   3045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Left            =   0
      MousePointer    =   2  'Cross
      Picture         =   "frmSetBgCol.frx":0000
      ScaleHeight     =   1200
      ScaleWidth      =   3045
      TabIndex        =   0
      Tag             =   "tab1"
      Top             =   0
      Width           =   3045
   End
End
Attribute VB_Name = "frmSetBgCol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Picture2_MouseUp(button As Integer, shift As Integer, X As Single, Y As Single)

    If X > (1 * Screen.TwipsPerPixelX) And Y > (1 * Screen.TwipsPerPixelY) And X < Picture2.Width - (1 * Screen.TwipsPerPixelX) And Y < Picture2.Height - (1 * Screen.TwipsPerPixelY) Then

        frmMain.BgColorRGB = Picture2.Point(X, Y)
    
        Unload Me

    End If

End Sub
