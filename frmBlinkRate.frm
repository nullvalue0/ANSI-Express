VERSION 5.00
Begin VB.Form frmBlinkRate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Blink Rate"
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1905
   Icon            =   "frmBlinkRate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   1905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtRate 
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "(ms on, ms off)"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Blink Rate:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmBlinkRate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    SaveSetting "ASNIExpress", "Settings", "BlinkRate", txtRate.Text
    frmMain.tmrBlink.Interval = txtRate.Text
    Unload Me
End Sub

Private Sub Form_Load()
    txtRate.Text = frmMain.tmrBlink.Interval
End Sub
