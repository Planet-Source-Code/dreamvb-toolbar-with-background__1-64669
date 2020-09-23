VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmMain 
   Caption         =   "Toolbar Background"
   ClientHeight    =   2970
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   ScaleHeight     =   2970
   ScaleWidth      =   6660
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      Height          =   405
      Left            =   60
      TabIndex        =   3
      Top             =   2280
      Width           =   2085
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Set Toolbar BackGround"
      Height          =   405
      Left            =   60
      TabIndex        =   2
      Top             =   1710
      Width           =   2085
   End
   Begin VB.CommandButton cmdBk 
      Caption         =   "Set Toolbar Backcolor"
      Height          =   405
      Left            =   45
      TabIndex        =   1
      Top             =   1125
      Width           =   2085
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4785
      Top             =   975
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B52
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":36A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":51F6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   840
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6660
      _ExtentX        =   11748
      _ExtentY        =   1482
      ButtonWidth     =   1455
      ButtonHeight    =   1429
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   1200
      Left            =   3150
      Picture         =   "frmMain.frx":6D48
      Top             =   1350
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   1
      X1              =   0
      X2              =   825
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   15
      X2              =   840
      Y1              =   855
      Y2              =   855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBk_Click()
    Call SetToolbarBK(Toolbar1.hwnd, vbYellow)
End Sub

Private Sub cmdexit_Click()
    MsgBox "Toolbar Extander by DreamVB.", vbInformation, "Exit.."
    Unload frmMain
End Sub

Private Sub Command1_Click()
    SetToolbarBG Toolbar1.hwnd, Image1.Picture
End Sub

Private Sub Form_Resize()
    Line1(0).X2 = frmMain.ScaleWidth
    Line1(1).X2 = Line1(0).X2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SetToolbarBK Toolbar1.hwnd, vbButtonFace
    Set frmMain = Nothing
End Sub
