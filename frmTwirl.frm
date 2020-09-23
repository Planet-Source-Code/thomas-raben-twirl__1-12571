VERSION 5.00
Begin VB.Form frmTwirl 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Twirl - By E1 - Thomas Raben"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   2325
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Save"
      Height          =   255
      Left            =   1500
      TabIndex        =   8
      Top             =   2940
      Width           =   795
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Browse"
      Height          =   255
      Left            =   660
      TabIndex        =   7
      Top             =   2940
      Width           =   795
   End
   Begin VB.PictureBox ProgressB 
      Height          =   255
      Left            =   60
      ScaleHeight     =   195
      ScaleWidth      =   2175
      TabIndex        =   5
      Top             =   2340
      Visible         =   0   'False
      Width           =   2235
      Begin VB.PictureBox Progress 
         BackColor       =   &H80000002&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   0
         ScaleHeight     =   195
         ScaleWidth      =   375
         TabIndex        =   6
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Draw"
      Height          =   255
      Left            =   1500
      TabIndex        =   4
      Top             =   2640
      Width           =   795
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   50
      Left            =   60
      Max             =   200
      TabIndex        =   1
      Top             =   2340
      Value           =   100
      Width           =   2235
   End
   Begin VB.PictureBox TwirlPic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      Height          =   2235
      Left            =   60
      ScaleHeight     =   145
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   145
      TabIndex        =   0
      Top             =   60
      Width           =   2235
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Twirl Value:"
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   2700
      Width           =   1215
   End
End
Attribute VB_Name = "frmTwirl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long


Private Sub Command1_Click()

    Render_Twirl Me.HScroll1.Value - (Me.HScroll1.Max / 2)
End Sub

Private Sub Command2_Click()
    Me.Enabled = False

    frmLoad.Show
    
End Sub

Private Sub Command3_Click()
    Me.Enabled = False
    frmSave.Show
    
End Sub

Private Sub Form_Load()
    Draw_Twirl 0
    frmImage.Show
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
    
End Sub

Private Sub HScroll1_Change()
    Draw_Twirl Me.HScroll1.Value - (Me.HScroll1.Max / 2)
    Me.Label1.Caption = Me.HScroll1.Value - (Me.HScroll1.Max / 2)
    Me.TwirlPic.SetFocus
End Sub

Private Sub HScroll1_Scroll()
    Draw_Twirl Me.HScroll1.Value - (Me.HScroll1.Max / 2)
    Me.Label1.Caption = Me.HScroll1.Value - (Me.HScroll1.Max / 2)
    Me.TwirlPic.SetFocus
End Sub


'DRAW TWIRL
Public Sub Draw_Twirl(Angle As Double)
    Dim Rad As Double
    Dim A As Double
    Dim B As Double
    Dim x As Double
    Dim y As Double
    
    x = Me.TwirlPic.ScaleWidth / 2
    y = Me.TwirlPic.ScaleHeight / 2
    
    B = Angle / 10000
    
    Me.TwirlPic.Cls
    
    For Rad = 100 To 0 Step -0.1
        A = A + B
        SetPixelV Me.TwirlPic.hdc, x + Cos(A) * Rad, y + Sin(A) * Rad, 0
        SetPixelV Me.TwirlPic.hdc, x - Cos(A) * Rad, y - Sin(A) * Rad, 0
        SetPixelV Me.TwirlPic.hdc, x - Sin(A) * Rad, y + Cos(A) * Rad, 0
        SetPixelV Me.TwirlPic.hdc, x + Sin(A) * Rad, y - Cos(A) * Rad, 0
    Next Rad
    
    
    
End Sub

Public Sub Render_Twirl(Angle As Double)
    Dim Rad As Double
    Dim A As Double
    Dim B As Double
    Dim x As Double
    Dim y As Double
    Dim R As Double
    Dim C As Long

    Dim OS_Y As Integer
    
    Const PI = 3.1415
    
    x = frmImage.Source.ScaleWidth / 2
    y = frmImage.Source.ScaleHeight / 2
    
    B = Angle / ((frmImage.Source.ScaleWidth / 2) * 100)
    
    frmImage.Dest.Picture = frmImage.Source.Picture
    frmImage.Dest.Visible = True
    
    Me.HScroll1.Enabled = False
    Me.Command1.Enabled = False
    Me.Command2.Enabled = False
    Me.Command3.Enabled = False
    Me.ProgressB.Visible = True
    Me.Progress.Width = 1
    
    For Rad = y To 0 Step -0.1
        A = A + B
        For R = 0 To PI * 2 Step (frmImage.Source.ScaleWidth / 2) / ((frmImage.Source.ScaleWidth / 2) * 100)
            C = GetPixel(frmImage.Source.hdc, (x + Cos(R) * Rad), (y + Sin(R) * Rad))
            SetPixelV frmImage.Dest.hdc, (x + Cos(A + R) * Rad), (y + Sin(A + R) * Rad), C
        Next R
        Me.Progress.Width = Me.ProgressB.ScaleWidth / 100 * (100 - (Rad / y * 100))
        frmImage.Dest.Refresh
        DoEvents
    Next Rad

    Me.HScroll1.Enabled = True
    Me.Command1.Enabled = True
    Me.Command2.Enabled = True
    Me.Command3.Enabled = True
    Me.ProgressB.Visible = False
    frmImage.Show
    
    frmImage.Dest.Refresh
    frmImage.Source.Refresh

End Sub

