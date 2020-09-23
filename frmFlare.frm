VERSION 5.00
Begin VB.Form frmFlare 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Flare"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Source 
      Height          =   3060
      Left            =   3240
      Picture         =   "frmFlare.frx":0000
      ScaleHeight     =   3000
      ScaleWidth      =   3000
      TabIndex        =   4
      Top             =   60
      Width           =   3060
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Draw"
      Height          =   315
      Left            =   1980
      TabIndex        =   3
      Top             =   3780
      Width           =   1155
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   33
      Left            =   60
      Max             =   99
      Min             =   3
      TabIndex        =   1
      Top             =   3480
      Value           =   30
      Width           =   3075
   End
   Begin VB.PictureBox Dest 
      AutoRedraw      =   -1  'True
      Height          =   3060
      Left            =   60
      MousePointer    =   2  'Cross
      Picture         =   "frmFlare.frx":1D504
      ScaleHeight     =   200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   0
      Top             =   60
      Width           =   3060
      Begin VB.Shape FlarePos 
         DrawMode        =   6  'Mask Pen Not
         Height          =   450
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   450
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Flare Size:"
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   3180
      Width           =   2115
   End
End
Attribute VB_Name = "frmFlare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type zRGB
    R As Long
    G As Long
    B As Long
End Type


Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long


Dim SetFlare As Boolean
Dim Red As Integer
Dim Green As Integer
Dim Blue As Integer


Private Sub Command1_Click()
    DrawFlare Int(Me.FlarePos.Left + Me.FlarePos.Width / 2), Int(Me.FlarePos.Top + Me.FlarePos.Height / 2), CDbl(Me.HScroll1.Value / 2)
    
End Sub

Private Sub Dest_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    SetFlare = True
    Me.FlarePos.Move x - Me.FlarePos.Width / 2, y - Me.FlarePos.Height / 2
    Me.FlarePos.Visible = True
    
End Sub

Private Sub Dest_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If SetFlare = True Then
        Me.FlarePos.Move x - Me.FlarePos.Width / 2, y - Me.FlarePos.Height / 2
        Me.FlarePos.Visible = True
    End If
End Sub

Private Sub Dest_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.FlarePos.Move x - Me.FlarePos.Width / 2, y - Me.FlarePos.Height / 2
    SetFlare = False
    Me.FlarePos.Visible = True
    
End Sub

Private Sub HScroll1_Change()
    Me.FlarePos.Width = Me.HScroll1.Value
    Me.FlarePos.Height = Me.HScroll1.Value
    Me.Dest.SetFocus
    Me.FlarePos.Visible = True
End Sub

Private Sub HScroll1_Scroll()
    Me.FlarePos.Width = Me.HScroll1.Value
    Me.FlarePos.Height = Me.HScroll1.Value
    Me.Dest.SetFocus
    Me.FlarePos.Visible = True
End Sub

Private Sub DrawFlare(x As Integer, y As Integer, Rad As Double)
    Dim i As Double
    Dim R As Double
    Dim C As Long
    Dim Rc As zRGB
    
    Const PI = 3.1415
    
    Me.Command1.Enabled = False
    'Me.Dest.Picture = Me.Source.Picture
    
    For i = 0 To PI * 2 Step 0.01
        For R = 0 To Rad
            C = GetPixel(Me.Source.hdc, x + Cos(i) * R, y + Sin(i) * R)
            'Debug.Print C
            'Get_RGB RGB(200, 0, 200)
            Rc = LongToRGB(C)
            
            On Error Resume Next
            Red = Rc.R + 10 * ((R + 1))
            Green = Rc.G + 10 * ((R + 1))
            Blue = Rc.B + 10 * ((R + 1))
            
            If Red > 255 Then Red = 255
            If Green > 255 Then Green = 255
            If Blue > 255 Then Blue = 255
            
            
            'Debug.Print Red & " " & Green & " " & Blue
            SetPixelV Me.Dest.hdc, x + Cos(i) * R, y + Sin(i) * R, RGB(Red, Green, Blue)
        Next R
    Next i
    Me.Dest.Refresh
    
    Me.Command1.Enabled = True
    Me.FlarePos.Visible = False
End Sub




Private Function LongToRGB(ColorValue As Long) As zRGB
    Dim rCol As Long, gCol As Long, bCol As Long
    rCol = ColorValue And &H10000FF 'this uses binary comparason
    gCol = (ColorValue And &H100FF00) / (2 ^ 8)
    bCol = (ColorValue And &H1FF0000) / (2 ^ 16)
    LongToRGB.R = rCol
    LongToRGB.G = gCol
    LongToRGB.B = bCol
End Function
