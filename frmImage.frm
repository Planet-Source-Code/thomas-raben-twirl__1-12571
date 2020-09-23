VERSION 5.00
Begin VB.Form frmImage 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Image"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   4020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Dest 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   3060
      Left            =   0
      ScaleHeight     =   200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3060
   End
   Begin VB.PictureBox Source 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3060
      Left            =   0
      OLEDropMode     =   1  'Manual
      Picture         =   "frmImage.frx":0000
      ScaleHeight     =   200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   1
      Top             =   0
      Width           =   3060
   End
End
Attribute VB_Name = "frmImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    frmImage.Width = frmImage.Source.Width + (frmImage.Width - frmImage.ScaleWidth)
    frmImage.Height = frmImage.Source.Height + (frmImage.Height - frmImage.ScaleHeight)

End Sub

