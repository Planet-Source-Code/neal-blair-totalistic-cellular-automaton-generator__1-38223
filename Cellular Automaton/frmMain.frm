VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cellular Automaton"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   177
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   345
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdProcess 
      Caption         =   "&Process"
      Height          =   375
      Left            =   3960
      TabIndex        =   0
      Top             =   1800
      Width           =   855
   End
   Begin VB.ComboBox cmbRS 
      Height          =   315
      Index           =   4
      ItemData        =   "frmMain.frx":0000
      Left            =   2880
      List            =   "frmMain.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   960
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Caption         =   "Initial Constants:"
      Height          =   2415
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   4935
      Begin MSComCtl2.UpDown udIterations 
         Height          =   255
         Left            =   1800
         TabIndex        =   23
         Top             =   1560
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         _Version        =   393216
         Value           =   250
         Increment       =   25
         Max             =   9975
         Min             =   25
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udZoomFactor 
         Height          =   255
         Left            =   1680
         TabIndex        =   22
         Top             =   1200
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         _Version        =   393216
         Value           =   10
         Max             =   31
         Min             =   2
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtZoomFactor 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   20
         Text            =   "10"
         Top             =   1200
         Width           =   375
      End
      Begin VB.ComboBox cmbPrintmode 
         Height          =   315
         ItemData        =   "frmMain.frx":001A
         Left            =   1320
         List            =   "frmMain.frx":0027
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox txtIterations 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   14
         Text            =   "250"
         Top             =   1560
         Width           =   495
      End
      Begin VB.ComboBox cmbRS 
         Height          =   315
         Index           =   2
         ItemData        =   "frmMain.frx":0040
         Left            =   3720
         List            =   "frmMain.frx":004D
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   840
         Width           =   495
      End
      Begin VB.ComboBox cmbRS 
         Height          =   315
         Index           =   3
         ItemData        =   "frmMain.frx":005A
         Left            =   3240
         List            =   "frmMain.frx":0067
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   840
         Width           =   495
      End
      Begin VB.ComboBox cmbRS 
         Height          =   315
         Index           =   5
         ItemData        =   "frmMain.frx":0074
         Left            =   2280
         List            =   "frmMain.frx":0081
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   840
         Width           =   495
      End
      Begin VB.ComboBox cmbRS 
         Height          =   315
         Index           =   6
         ItemData        =   "frmMain.frx":008E
         Left            =   1800
         List            =   "frmMain.frx":009B
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   840
         Width           =   495
      End
      Begin VB.ComboBox cmbRS 
         Height          =   315
         Index           =   1
         ItemData        =   "frmMain.frx":00A8
         Left            =   4200
         List            =   "frmMain.frx":00B5
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   840
         Width           =   495
      End
      Begin VB.ComboBox cmbRS 
         Height          =   315
         Index           =   7
         ItemData        =   "frmMain.frx":00C2
         Left            =   1320
         List            =   "frmMain.frx":00CF
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   840
         Width           =   495
      End
      Begin VB.ComboBox cmbSeed 
         Height          =   315
         ItemData        =   "frmMain.frx":00DC
         Left            =   1320
         List            =   "frmMain.frx":00E9
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Average Color:"
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   600
         Width           =   1050
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   3720
         TabIndex        =   27
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   3240
         TabIndex        =   26
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label7 
         BackColor       =   &H00404040&
         Height          =   255
         Left            =   1800
         TabIndex        =   25
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000001&
         Height          =   255
         Left            =   2280
         TabIndex        =   24
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Block Size:"
         Height          =   195
         Left            =   420
         TabIndex        =   21
         Top             =   1260
         Width           =   795
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "by Neal Blair"
         Height          =   195
         Left            =   3960
         TabIndex        =   19
         Top             =   120
         Width           =   885
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "0"
         Height          =   255
         Left            =   4200
         TabIndex        =   18
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "1"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2760
         TabIndex        =   17
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "2"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1320
         TabIndex        =   16
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblPrintMode 
         AutoSize        =   -1  'True
         Caption         =   "Print Mode:"
         Height          =   195
         Left            =   420
         TabIndex        =   13
         Top             =   1980
         Width           =   810
      End
      Begin VB.Label lblIterations 
         AutoSize        =   -1  'True
         Caption         =   "Iterations:"
         Height          =   195
         Left            =   540
         TabIndex        =   12
         Top             =   1620
         Width           =   690
      End
      Begin VB.Label lblReplaceString 
         AutoSize        =   -1  'True
         Caption         =   "Replace With:"
         Height          =   195
         Left            =   140
         TabIndex        =   11
         Top             =   900
         Width           =   1020
      End
      Begin VB.Label lblSeed 
         AutoSize        =   -1  'True
         Caption         =   "Seed:"
         Height          =   195
         Left            =   800
         TabIndex        =   10
         Top             =   300
         Width           =   420
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdProcess_Click()
If Trim(txtIterations.Text) = "" Then Exit Sub
Unload frmDisplay
cmdProcess.Enabled = False


SEED = ValueToColor(cmbSeed.Text)
ITERATIONS = txtIterations.Text
PRINTMODE = cmbPrintmode.Text
ReplaceString = cmbRS(7).Text & cmbRS(6).Text & cmbRS(5).Text & cmbRS(4).Text & cmbRS(3).Text & cmbRS(2).Text & cmbRS(1).Text
ZOOMFACTOR = txtZoomFactor.Text

ReplaceColor(1) = ValueToColor(cmbRS(1).Text)
ReplaceColor(2) = ValueToColor(cmbRS(2).Text)
ReplaceColor(3) = ValueToColor(cmbRS(3).Text)
ReplaceColor(4) = ValueToColor(cmbRS(4).Text)
ReplaceColor(5) = ValueToColor(cmbRS(5).Text)
ReplaceColor(6) = ValueToColor(cmbRS(6).Text)
ReplaceColor(7) = ValueToColor(cmbRS(7).Text)

MakeString ITERATIONS

LineString(1) = SEED

Dim IterationCount As Integer

For IterationCount = 1 To ITERATIONS - 1
    frmMain.Caption = "Processing: " & Round(IterationCount / (ITERATIONS - 1) * 100, 1) & "%"
    LineString(IterationCount + 1) = ProcessString(IterationCount)
Next IterationCount

Select Case PRINTMODE
    Case "Text":
        PrintAutomaton "Text"
    Case "HTML":
        PrintAutomaton "HTML"
    Case "Display":
        Load frmDisplay
        frmDisplay.Show
        frmDisplay.Cls
        frmDisplay.hsbDisplay.Value = frmDisplay.hsbDisplay.Max / 2
        frmDisplay.vsbDisplay.Value = 1
End Select

Me.Caption = "Done!"

cmdProcess.Enabled = True
End Sub

Private Sub Form_Load()

cmbSeed.ListIndex = 2

cmbRS(1).ListIndex = 0
cmbRS(2).ListIndex = 1
cmbRS(3).ListIndex = 2
cmbRS(4).ListIndex = 1
cmbRS(5).ListIndex = 0
cmbRS(6).ListIndex = 0
cmbRS(7).ListIndex = 1

cmbPrintmode.ListIndex = 2

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload frmDisplay
End Sub

Private Sub udIterations_Change()
txtIterations.Text = udIterations.Value
End Sub

Private Sub udZoomFactor_Change()
txtZoomFactor.Text = udZoomFactor.Value
End Sub
