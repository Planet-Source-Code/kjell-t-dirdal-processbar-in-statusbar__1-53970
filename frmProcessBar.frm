VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   6915
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8955
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar sbMain 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   6600
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   0
      ScaleHeight     =   1215
      ScaleWidth      =   8955
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   8955
      Begin MSComctlLib.ProgressBar pbMain 
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'SetParent
Private Declare Function SetParent Lib "user32.dll" _
    (ByVal hWndChild As Long, _
ByVal hWndNewParent As Long) As Long
'SendMessage
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" _
    (ByVal hWnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long

'Type RECT
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'Const
Private Const WM_USER As Long = &H400
Private Const SB_GETRECT As Long = (WM_USER + 10)

Private Sub MDIForm_Load()
    Dim rectPanel As RECT
    'Read panel coordinates and dimensions (pixel scale)
    '1 means the second panel of status bar (0-index based)
    SendMessage sbMain.hWnd, SB_GETRECT, 1, rectPanel
    'Transform coordinates from pixel to twip
    'Bottom contains the height, Right contains the width
    rectPanel.Top = (rectPanel.Top * Screen.TwipsPerPixelY)
    rectPanel.Left = (rectPanel.Left * Screen.TwipsPerPixelX)
    rectPanel.Bottom = (rectPanel.Bottom * Screen.TwipsPerPixelY) - rectPanel.Top
    rectPanel.Right = (rectPanel.Right * Screen.TwipsPerPixelX) - rectPanel.Left
    'Move progress bar inside the statusbar panel
    SetParent pbMain.hWnd, sbMain.hWnd
    pbMain.Move rectPanel.Left, rectPanel.Top, rectPanel.Right, rectPanel.Bottom
End Sub

'====================================================
'Sub: ShowStatusBarPercent
'====================================================
Public Sub ShowStatusBarPercent(ByVal iPercent_ As Integer)
    'The progress bar as default accepts only values from 0 to 100
    'This check is necessary to avoid errors
    If iPercent_ > pbMain.Max Then
        iPercent_ = pbMain.Max
    ElseIf iPercent_ < pbMain.Min Then
        iPercent_ = pbMain.Min
    End If
    pbMain.Value = iPercent_
End Sub

'====================================================
'Sub: MyExample
'====================================================
Public Sub MyExample()
    Dim n As Integer

    For n = 1 To 10000
        ShowStatusBarPercent n / 100
        DoEvents
        'Do something
    Next n
    ShowStatusBarPercent 0
End Sub

Private Sub sbMain_PanelClick(ByVal Panel As MSComctlLib.Panel)
    MyExample
End Sub
