VERSION 5.00
Begin VB.Form frmScreenSize 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Screen Size"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1935
   Icon            =   "frmScreenSize.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   1935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox txtRow 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox txtCol 
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Rows:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Colums:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmScreenSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    SaveSetting "ASNIExpress", "Settings", "xScreenSize", txtCol.Text
    SaveSetting "ASNIExpress", "Settings", "yScreenSize", txtRow.Text
    frmMain.xScreenSize = txtCol.Text
    frmMain.yScreenSize = txtRow.Text
    Unload Me
End Sub

Private Sub Form_Load()
    txtCol.Text = frmMain.xScreenSize
    txtRow.Text = frmMain.yScreenSize
End Sub

