VERSION 5.00
Begin VB.Form frmDisplayRate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Display Rate"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1935
   Icon            =   "frmDisplayRate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   1935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtDispRate 
      Height          =   285
      Left            =   960
      TabIndex        =   2
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
      TabIndex        =   0
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0 = unlimited"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Enter baud rate to emulate while loading screen"
      Height          =   615
      Left            =   0
      TabIndex        =   4
      Top             =   450
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Baud Rate:"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmDisplayRate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    For i = 0 To 9
        frmMain.mnuSpeed(i).Checked = False
    Next i
    frmMain.mnuSpeed(8).Checked = True
    SaveSetting "ASNIExpress", "Settings", "DisplayRate", txtDispRate.Text
    frmMain.DisplayRate = txtDispRate.Text
    Unload Me
End Sub

Private Sub Form_Load()
    txtDispRate.Text = frmMain.DisplayRate
End Sub
