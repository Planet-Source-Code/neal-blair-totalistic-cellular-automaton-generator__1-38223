VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDisplay 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Cellular Automaton"
   ClientHeight    =   9030
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12015
   LinkTopic       =   "Form1"
   ScaleHeight     =   602
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   801
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSB 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   255
      Left            =   11760
      TabIndex        =   5
      Top             =   8760
      Width           =   255
   End
   Begin MSComCtl2.FlatScrollBar vsbDisplay 
      Height          =   8775
      Left            =   11760
      TabIndex        =   4
      Top             =   0
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   15478
      _Version        =   393216
      LargeChange     =   5
      Min             =   1
      Orientation     =   1179648
      Value           =   1
   End
   Begin MSComCtl2.FlatScrollBar hsbDisplay 
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   8760
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      LargeChange     =   10
      Min             =   1
      Orientation     =   1179649
      Value           =   1
   End
   Begin VB.PictureBox Color 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   7
      Left            =   0
      Picture         =   "frmDisplay.frx":0000
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox Color 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   4
      Left            =   0
      Picture         =   "frmDisplay.frx":0C44
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox Color 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   1
      Left            =   0
      Picture         =   "frmDisplay.frx":1888
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
cmdSB.Top = Me.ScaleHeight - cmdSB.Height
cmdSB.Left = Me.ScaleWidth - cmdSB.Width
hsbDisplay.Width = Me.ScaleWidth - vsbDisplay.Width
vsbDisplay.Height = Me.ScaleHeight - hsbDisplay.Height
hsbDisplay.Top = Me.ScaleHeight - hsbDisplay.Height
vsbDisplay.Left = Me.ScaleWidth - vsbDisplay.Width

DisplayX = (Me.ScaleWidth - vsbDisplay.Width) / ZOOMFACTOR
DisplayY = (Me.ScaleHeight - hsbDisplay.Height) / ZOOMFACTOR

If DisplayY > ITERATIONS Then DisplayY = ITERATIONS

vsbDisplay.Max = ITERATIONS - DisplayY
hsbDisplay.Max = Len(LineString(ITERATIONS)) - DisplayX
Me.Caption = "H:" & hsbDisplay.Value & " V:" & vsbDisplay.Value

GenerateWindow
End Sub

Private Sub hsbDisplay_Change()
Me.Caption = "H:" & hsbDisplay.Value & " V:" & vsbDisplay.Value
GenerateWindow
End Sub

Private Sub hsbDisplay_Scroll()
Me.Caption = "H:" & hsbDisplay.Value & " V:" & vsbDisplay.Value
GenerateWindow
End Sub

Private Sub vsbDisplay_Change()
Me.Caption = "H:" & hsbDisplay.Value & " V:" & vsbDisplay.Value
GenerateWindow
End Sub

Private Sub vsbDisplay_Scroll()
Me.Caption = "H:" & hsbDisplay.Value & " V:" & vsbDisplay.Value
GenerateWindow
End Sub
