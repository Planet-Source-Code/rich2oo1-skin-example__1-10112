VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Skin Example by rich"
   ClientHeight    =   1605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4380
   LinkTopic       =   "Form1"
   ScaleHeight     =   107
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   292
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picBut4 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   2280
      ScaleHeight     =   360
      ScaleWidth      =   1155
      TabIndex        =   7
      Top             =   1080
      Width           =   1155
   End
   Begin VB.PictureBox picBut3 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   2280
      ScaleHeight     =   360
      ScaleWidth      =   1155
      TabIndex        =   6
      Top             =   600
      Width           =   1155
   End
   Begin VB.PictureBox picBut2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   120
      ScaleHeight     =   360
      ScaleWidth      =   1155
      TabIndex        =   5
      Top             =   1080
      Width           =   1155
   End
   Begin VB.PictureBox picBut1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   120
      ScaleHeight     =   360
      ScaleWidth      =   1155
      TabIndex        =   4
      Top             =   600
      Width           =   1155
   End
   Begin VB.PictureBox picClose 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   4020
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   3
      Top             =   60
      Width           =   315
   End
   Begin VB.PictureBox picMin 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   3660
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   2
      Top             =   60
      Width           =   315
   End
   Begin VB.PictureBox picTitle 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   60
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   233
      TabIndex        =   1
      Top             =   60
      Width           =   3495
   End
   Begin VB.PictureBox picSkin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1620
      Left            =   120
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   1620
      ScaleWidth      =   4860
      TabIndex        =   0
      Top             =   2160
      Width           =   4860
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Skin Example by rich
'---------------
'I made this example for all of you people
'who have no clue how to use the bitblt statement.
'I have just recently really learned how to use it
'and i wanted to share my knowledge. Here is some basic
'info on the BitBlt statement. Also, please dont critisize
'my artistic abilities. i am not that good with photoshop
'yet.
'---------------
'the declare for bitblt is:
'Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'
'here is a description of all of the arguments:

'ByVal hDestDC As Long
'Destination for the extracted part of the skin

'ByVal X As Long
'X to begin at. Just put 0

'ByVal Y As Long
'Y to begin at. Just put 0

'ByVal nWidth As Long
'Width of the destination

'ByVal nHeight As Long
'height of the destination

'ByVal hSrcDC As Long
'Skin's HDC. ex : picSkin.hdc

'ByVal xSrc As Long
'X of the part of the skin being extracted

'ByVal ySrc As Long
'Y of the part of the skin being extracted

'ByVal dwRop As Long
'Raster Operation Code. I use the SRCCopy Const


Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Const SRCCOPY = &HCC0020

Private Sub Form_Load()
Call BitBlt(picTitle.hdc, 0, 0, 233, 20, picSkin.hdc, 2, 0, SRCCOPY)
Call BitBlt(picMin.hdc, 0, 0, 20, 15, picSkin.hdc, 235, 0, SRCCOPY)
Call BitBlt(picClose.hdc, 0, 0, 20, 15, picSkin.hdc, 258, 0, SRCCOPY)
Call BitBlt(picBut1.hdc, 0, 0, 77, 24, picSkin.hdc, 0, 25, SRCCOPY)
Call BitBlt(picBut2.hdc, 0, 0, 77, 24, picSkin.hdc, 0, 54, SRCCOPY)
Call BitBlt(picBut3.hdc, 0, 0, 77, 24, picSkin.hdc, 165, 22, SRCCOPY)
Call BitBlt(picBut4.hdc, 0, 0, 77, 24, picSkin.hdc, 166, 52, SRCCOPY)
picTitle.Refresh
picMin.Refresh
picClose.Refresh
picBut1.Refresh
picBut2.Refresh
picBut3.Refresh
picBut4.Refresh
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Note: ALL PICTURE BOXES AUTO REDRAW FEATURES MUST
'BE SET TO TRUE OR IT WONT WORK!!
Call BitBlt(picTitle.hdc, 0, 0, 233, 20, picSkin.hdc, 2, 0, SRCCOPY)
Call BitBlt(picMin.hdc, 0, 0, 20, 15, picSkin.hdc, 235, 0, SRCCOPY)
Call BitBlt(picClose.hdc, 0, 0, 20, 15, picSkin.hdc, 258, 0, SRCCOPY)
Call BitBlt(picBut1.hdc, 0, 0, 77, 24, picSkin.hdc, 0, 25, SRCCOPY)
Call BitBlt(picBut2.hdc, 0, 0, 77, 24, picSkin.hdc, 0, 54, SRCCOPY)
Call BitBlt(picBut3.hdc, 0, 0, 77, 24, picSkin.hdc, 165, 22, SRCCOPY)
Call BitBlt(picBut4.hdc, 0, 0, 77, 24, picSkin.hdc, 166, 52, SRCCOPY)
picTitle.Refresh
picMin.Refresh
picClose.Refresh
picBut1.Refresh
picBut2.Refresh
picBut3.Refresh
picBut4.Refresh
End Sub

Private Sub picBut1_Click()
MsgBox "You clicked button 1", , "Button 1"
End Sub

Private Sub picBut1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call BitBlt(picBut1.hdc, 0, 0, 77, 24, picSkin.hdc, 82, 23, SRCCOPY)
picBut1.Refresh
End Sub

Private Sub picBut2_Click()
MsgBox "You clicked button 2", , "Button 2"
End Sub

Private Sub picBut2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call BitBlt(picBut2.hdc, 0, 0, 77, 24, picSkin.hdc, 85, 52, SRCCOPY)
picBut2.Refresh
End Sub

Private Sub picBut3_Click()
MsgBox "You clicked button 3", , "Button 3"
End Sub

Private Sub picBut3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call BitBlt(picBut3.hdc, 0, 0, 77, 24, picSkin.hdc, 246, 22, SRCCOPY)
picBut3.Refresh
End Sub

Private Sub picBut4_Click()
MsgBox "You clicked button 4", , "Button 4"
End Sub

Private Sub picBut4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call BitBlt(picBut4.hdc, 0, 0, 77, 24, picSkin.hdc, 245, 52, SRCCOPY)
picBut4.Refresh
End Sub

Private Sub picClose_Click()
Unload Me
End Sub

Private Sub picClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call BitBlt(picClose.hdc, 0, 0, 20, 15, picSkin.hdc, 303, 0, SRCCOPY)
picClose.Refresh
End Sub

Private Sub picMin_Click()
WindowState = 1
End Sub

Private Sub picMin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call BitBlt(picMin.hdc, 0, 0, 20, 15, picSkin.hdc, 282, 0, SRCCOPY)
picMin.Refresh
End Sub

Private Sub picTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call FormDrag(Me)
End Sub
