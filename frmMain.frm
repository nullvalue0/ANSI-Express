VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "ANSI Express"
   ClientHeight    =   4500
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9600
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4500
   ScaleWidth      =   9600
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrBlink 
      Interval        =   200
      Left            =   480
      Top             =   1200
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   480
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Open ANSI"
      Filter          =   "ANSI Files (*.ans;*.asc; *.txt; *.vt)|*.ANS;*.ASC;*.TXT;*.VT|All Files (*.*)|*.*"
   End
   Begin VB.PictureBox picNorm 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4500
      Left            =   0
      ScaleHeight     =   300
      ScaleMode       =   0  'User
      ScaleWidth      =   640
      TabIndex        =   0
      Top             =   0
      Width           =   9600
   End
   Begin VB.PictureBox picBlnk 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4500
      Left            =   0
      ScaleHeight     =   300
      ScaleMode       =   0  'User
      ScaleWidth      =   640
      TabIndex        =   1
      Top             =   0
      Width           =   9600
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open ANSI..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuDash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy Screen to Clipboard"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuDash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuDash3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&Tools"
      Begin VB.Menu mnuDraw 
         Caption         =   "Draw &Speed"
         Begin VB.Menu mnuSpeed 
            Caption         =   "300 baud"
            Index           =   0
         End
         Begin VB.Menu mnuSpeed 
            Caption         =   "1200 baud"
            Index           =   1
         End
         Begin VB.Menu mnuSpeed 
            Caption         =   "2400 baud"
            Index           =   2
         End
         Begin VB.Menu mnuSpeed 
            Caption         =   "9600 baud"
            Index           =   3
         End
         Begin VB.Menu mnuSpeed 
            Caption         =   "14400 baud"
            Index           =   4
         End
         Begin VB.Menu mnuSpeed 
            Caption         =   "28800 baud"
            Index           =   5
         End
         Begin VB.Menu mnuSpeed 
            Caption         =   "33600 baud"
            Index           =   6
         End
         Begin VB.Menu mnuSpeed 
            Caption         =   "56700 baud"
            Index           =   7
         End
         Begin VB.Menu mnuSpeed 
            Caption         =   "Other..."
            Index           =   8
         End
         Begin VB.Menu mnuSpeed 
            Caption         =   "Unlimited"
            Index           =   9
         End
      End
      Begin VB.Menu mnuTerminal 
         Caption         =   "&Text Size"
         Begin VB.Menu mnuVerySmall 
            Caption         =   "Very Small"
         End
         Begin VB.Menu mnuSmall 
            Caption         =   "Small"
         End
         Begin VB.Menu mnuNormal 
            Caption         =   "Normal"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuLarge 
            Caption         =   "Large"
         End
         Begin VB.Menu mnuVeryLarge 
            Caption         =   "Very Large"
         End
      End
      Begin VB.Menu mnuAttrib 
         Caption         =   "&Text Attributes"
         Begin VB.Menu mnuBold 
            Caption         =   "&Bold"
         End
         Begin VB.Menu mnuItalic 
            Caption         =   "&Italic"
         End
      End
      Begin VB.Menu mnuDash5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuScreen 
         Caption         =   "&Screen Size..."
      End
      Begin VB.Menu mnuRate 
         Caption         =   "&Blink Rate..."
      End
      Begin VB.Menu mnuDash4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "&Clear Screen"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuRedraw 
         Caption         =   "&Redraw Screen"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuDash6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About ANSI Express..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'http://barney.cs.uni-potsdam.de/pipermail/kmud/1999-October/000008.html

'Something like:

' Set  to GetTickCount before starting.  Then before each character:

Dim x As Integer, Y As Integer, iLastColor As Long
Dim SaveX As Integer, SaveY As Integer, f As String
Dim xOffSet As Integer, yOffSet As Integer
Public xScreenSize As Integer, yScreenSize As Integer
Dim bFontBold As Boolean, bFontItalic As Boolean
Dim bBlinkUsed As Boolean
Dim MyTime As Double, TempTime As Double
Public DisplayRate As Long
Dim ScrollTop As Integer, ScrollBottom As Integer, iLines As Integer

Public Sub Display(data As String)
Dim c
Dim tx As Integer, ty As Integer
Dim bEscMode As Boolean
Dim bBold As Boolean
Dim bBlink As Boolean
Dim sCmd As String
Dim Commands() As String
MyTime = Timer
iLines = 0
ScrollTop = 1
ScrollBottom = yScreenSize
bEscMode = False
bBold = False
bBlink = False
bBlinkUsed = False
picBlnk.Visible = False
picNorm.Visible = True
sCmd = ""
    For i = 1 To Len(data)
        If DisplayRate > 0 Then
            TempTime = Timer  ' Don't know if you need ()
            If TempTime < MyTime Then
              DoEvents
              TempTime = Timer
              If (MyTime > TempTime + 0.005) Then
                Sleep 1000 * (MyTime - TempTime)
                DoEvents
              End If
            End If
            If DisplayRate > 0 Then MyTime = MyTime + (10# / DisplayRate)
        End If
        
        c = Mid(data, i, 1)
        If c = Chr(27) Then
            bEscMode = True
        ElseIf bEscMode = True Then
            If c <> "[" Then
                Select Case c
                Case "m"
                    Commands = Split(sCmd, ";")
                    sCmd = ""
                    For Each cmd In Commands
                        If cmd <> "" Then
                            Select Case cmd
                                Case 0
                                    picNorm.ForeColor = QBColor(7)
                                    picNorm.FillColor = QBColor(0)
                                    picBlnk.ForeColor = QBColor(7)
                                    picBlnk.FillColor = QBColor(0)
                                    iLastColor = 7
                                    bBold = False
                                    bBlink = False
                                Case 1
                                    picNorm.ForeColor = GetColor(iLastColor + 8)
                                    picBlnk.ForeColor = GetColor(iLastColor + 8)
                                    bBold = True
                                Case 5
                                    bBlink = True
                                    bBlinkUsed = True
                                Case 30 To 37
                                    If bBold = False Then
                                        picNorm.ForeColor = GetColor(cmd - 30)
                                        picBlnk.ForeColor = GetColor(cmd - 30)
                                    Else
                                        picNorm.ForeColor = GetColor((cmd - 30) + 8)
                                        picBlnk.ForeColor = GetColor(cmd - 30)
                                    End If
                                    iLastColor = cmd - 30
                                Case 40 To 47
                                    picNorm.FillColor = GetColor(cmd - 40)
                                    picBlnk.FillColor = GetColor(cmd - 40)
                            End Select
                        End If
                    Next
                    bEscMode = False
                Case "A"
                    If sCmd = "" Then
                        Y = Y - yOffSet
                    ElseIf IsNumeric(sCmd) = True Then
                        Y = Y - (sCmd * yOffSet)
                    End If
                    bEscMode = False
                    sCmd = ""
                Case "B"
                    If sCmd = "" Then
                        Y = Y + yOffSet
                        iLines = iLines + 1
                    ElseIf IsNumeric(sCmd) = True Then
                        Y = Y + (sCmd * yOffSet)
                        iLines = iLines + sCmd
                    End If
                    bEscMode = False
                    sCmd = ""
                Case "C"
                    If sCmd = "" Then
                        x = x + xOffSet
                    ElseIf IsNumeric(sCmd) = True Then
                        x = x + (sCmd * xOffSet)
                    End If
                    bEscMode = False
                    sCmd = ""
                Case "D"
                    If sCmd = "" Then
                        x = x - xOffSet
                    ElseIf IsNumeric(sCmd) = True Then
                        x = x - (sCmd * xOffSet)
                    End If
                    If x < 0 Then x = 0
                    bEscMode = False
                    sCmd = ""
                Case "L"
                    If sCmd = "" Then sCmd = 1
                    BitBlt picNorm.hDC, 0, Y + (yOffSet * sCmd), xOffSet * xScreenSize, (yOffSet * yScreenSize) * sCmd, picNorm.hDC, 0, Y, vbSrcCopy
                    BitBlt picBlnk.hDC, 0, Y + (yOffSet * sCmd), xOffSet * xScreenSize, (yOffSet * yScreenSize) * sCmd, picBlnk.hDC, 0, Y, vbSrcCopy
                    picNorm.Line (0, Y)-(xOffSet * xScreenSize, (Y + (yOffSet * sCmd)) - 1), picNorm.FillColor, BF
                    picBlnk.Line (0, Y)-(xOffSet * xScreenSize, (Y + (yOffSet * sCmd)) - 1), picBlnk.FillColor, BF
                    bEscMode = False
                    sCmd = ""
                Case "M", "Y"
                Case "H"
                    'sCmd = Replace(sCmd, "(", "")
                    If sCmd <> "" Then
                        l = InStr(1, sCmd, ";")
                        If l = 0 Then
                            Y = (sCmd - 1) * yOffSet
                            x = 0
                        ElseIf l = 1 Then
                            If sCmd = ";" Then
                                x = 0
                                Y = 0
                            Else
                                x = (Mid(sCmd, 2) - 1) * xOffSet
                                Y = 0
                            End If
                        Else
                            Y = CInt((Mid(sCmd, 1, l - 1)) - 1) * yOffSet
                            x = CInt((Mid(sCmd, l + 1)) - 1) * xOffSet
                        End If
                    Else
                        x = 0
                        Y = 0
                    End If
                    bEscMode = False
                    sCmd = ""
                Case "K"
                    If sCmd = "" Or sCmd = "0" Then
                        picNorm.Line (x, Y)-(xOffSet * xScreenSize, Y + yOffSet - 1), picNorm.FillColor, BF
                        picBlnk.Line (x, Y)-(xOffSet * xScreenSize, Y + yOffSet - 1), picNorm.FillColor, BF
                    ElseIf sCmd = 1 Then
                        picNorm.Line (0, Y)-(x, Y + yOffSet - 1), picNorm.FillColor, BF
                        picBlnk.Line (0, Y)-(x, Y + yOffSet - 1), picNorm.FillColor, BF
                    ElseIf sCmd = 2 Then
                        picNorm.Line (0, Y)-(xOffSet * xScreenSize, Y + yOffSet - 1), picNorm.FillColor, BF
                        picBlnk.Line (0, Y)-(xOffSet * xScreenSize, Y + yOffSet - 1), picNorm.FillColor, BF
                    End If
                Case "u"
                    x = SaveX
                    Y = SaveY
                    bEscMode = False
                    sCmd = ""
                Case "s"
                    SaveX = x
                    SaveY = Y
                    bEscMode = False
                    sCmd = ""
                Case "L"
                    scrollup 1, ScrollTop, ScrollBottom
                    bEscMode = False
                    sCmd = ""
                Case "M"
                    scrolldown 1, ScrollTop, ScrollBottom
                    bEscMode = False
                    sCmd = ""
                Case "J"
                    picNorm.Cls
                    picBlnk.Cls
                    x = 0
                    Y = 0
                    bEscMode = False
                    sCmd = ""
                Case "r"
                    If sCmd = "" Then
                        ScrollTop = 1
                        ScrollBottom = yScreenSize
                    Else
                        l = InStr(1, sCmd, ";")
                        If l = 0 Then
                            ScrollTop = sCmd
                            ScrollBottom = yScreenSize
                        Else
                            ScrollTop = Mid(sCmd, 1, l - 1)
                            ScrollBottom = Mid(sCmd, l + 1)
                        End If
                    End If
                    bEscMode = False
                    sCmd = ""
                Case "l", "h"
                    bEscMode = False
                    sCmd = ""
                Case Is > "?"
                    'MsgBox "unrecognized command: " & sCmd & c
                    bEscMode = False
                    sCmd = ""
                Case Else
                    sCmd = sCmd & c
                End Select
            End If
        ElseIf c = Chr(13) Or c = Chr(10) Then
            picNorm.CurrentX = x
            picNorm.CurrentY = Y
            picBlnk.CurrentX = x
            picBlnk.CurrentY = Y
            x = 0
            Y = Y + yOffSet
            iLines = iLines + 1
            If Y > ((yOffSet * ScrollBottom) - yOffSet) Then
                scrollup 1, ScrollTop, ScrollBottom
                Y = yOffSet * (ScrollBottom - 1)
            End If
        Else
            If x >= (xOffSet * xScreenSize) Then
                x = 0
                Y = Y + yOffSet
                iLines = iLines + 1
            End If
            
            If Y > ((yOffSet * ScrollBottom) - yOffSet) Then
                scrollup 1, ScrollTop, ScrollBottom
                Y = yOffSet * (ScrollBottom - 1)
            End If
            
            picNorm.Line (x, Y)-(x + (xOffSet - 1), Y + (yOffSet - 1)), picNorm.FillColor, BF
            picNorm.CurrentX = x
            picNorm.CurrentY = Y
            picNorm.Print c
            
            picBlnk.Line (x, Y)-(x + (xOffSet - 1), Y + (yOffSet - 1)), picBlnk.FillColor, BF
            picBlnk.CurrentX = x
            picBlnk.CurrentY = Y
            picBlnk.ForeColor = picNorm.ForeColor
            If bBlink = False Then picBlnk.Print c
            
            x = x + xOffSet
            
            tx = x / xOffSet
            ty = Y / yOffSet
        End If
    Next i
    If bBlinkUsed = True Then tmrBlink.Enabled = True
End Sub

Private Function GetColor(clr As Integer) As Long
    Select Case clr
        Case 0
            GetColor = QBColor(0)
        Case 1
            GetColor = QBColor(4)
        Case 2
            GetColor = QBColor(2)
        Case 3
            GetColor = QBColor(6)
        Case 4
            GetColor = QBColor(1)
        Case 5
            GetColor = QBColor(5)
        Case 6
            GetColor = QBColor(3)
        Case 7
            GetColor = QBColor(7)
        Case 8
            GetColor = QBColor(8)
        Case 9
            GetColor = QBColor(12)
        Case 10
            GetColor = QBColor(10)
        Case 11
            GetColor = QBColor(14)
        Case 12
            GetColor = QBColor(9)
        Case 13
            GetColor = QBColor(13)
        Case 14
            GetColor = QBColor(11)
        Case 15
            GetColor = QBColor(15)
    End Select
End Function

Private Sub Form_Load()
On Error GoTo Err_Handler
    RegApp
    picNorm.ForeColor = GetColor(7)
    picBlnk.ForeColor = GetColor(7)
    picNorm.FontSize = 9
    picBlnk.FontSize = 9
    xOffSet = GetSetting("ASNIExpress", "Settings", "xOffSet", 8)
    yOffSet = GetSetting("ASNIExpress", "Settings", "yOffSet", 12)
    xScreenSize = GetSetting("ASNIExpress", "Settings", "xScreenSize", 80)
    yScreenSize = GetSetting("ASNIExpress", "Settings", "yScreenSize", 24)
    ScrollTop = 1
    ScrollBottom = yScreenSize
    
    bFontBold = GetSetting("ASNIExpress", "Settings", "FontBold", False)
    mnuBold.Checked = bFontBold
    bFontItalic = GetSetting("ASNIExpress", "Settings", "FontItalic", False)
    mnuItalic.Checked = bFontItalic
    
    tmrBlink.Interval = GetSetting("ASNIExpress", "Settings", "BlinkRate", 200)
    DisplayRate = GetSetting("ASNIExpress", "Settings", "DisplayRate", 0)
    
    Select Case DisplayRate
        Case 300
            mnuSpeed(0).Checked = True
        Case 1200
            mnuSpeed(1).Checked = True
        Case 2400
            mnuSpeed(2).Checked = True
        Case 9600
            mnuSpeed(3).Checked = True
        Case 14400
            mnuSpeed(4).Checked = True
        Case 28800
            mnuSpeed(5).Checked = True
        Case 33600
            mnuSpeed(6).Checked = True
        Case 56700
            mnuSpeed(7).Checked = True
        Case 0
            mnuSpeed(9).Checked = True
        Case Else
            mnuSpeed(8).Checked = True
    End Select
    
    Select Case GetSetting("ASNIExpress", "Settings", "TextSize", 3)
        Case 1
            mnuVerySmall_Click
        Case 2
            mnuSmall_Click
        Case 3
            mnuNormal_Click
        Case 4
            mnuLarge_Click
        Case 5
            mnuVeryLarge_Click
    End Select
    
    If Command <> "" Then
        Me.Show
        DoEvents
        Dim t As String
        x = 0
        Y = 0
        picNorm.Cls
        picNorm.CurrentX = 0
        picNorm.CurrentY = 0
        picBlnk.Cls
        picBlnk.CurrentX = 0
        picBlnk.CurrentY = 0
        SaveX = 0
        SaveY = 0
        f = ""
        Open Command For Input As #1
        Do While Not EOF(1)
            Line Input #1, t
            f = f & t & vbCr
        Loop
        Close #1
        f = Left(f, Len(f) - 1)
        f = Replace(f, "[?7h", "")
        Display f
    End If
    
Exit Sub
Err_Handler:
        If Err.Number <> 32755 Then
            MsgBox Err.Description, vbCritical, "Error opening file"
        End If
End Sub

Private Sub mnuAbout_Click()
    MsgBox "ANSI Express version 1.2" & vbCr & vbCr & "(C) 2004-2008 Strinum Software" & vbCr & vbCr & "support@strinum.com" & vbCr & vbCr & "http://www.strinum.com", vbInformation, "About ANSI Express"
End Sub

Private Sub mnuBold_Click()
    If bFontBold = False Then
        bFontBold = True
        mnuBold.Checked = True
        SaveSetting "ASNIExpress", "Settings", "FontBold", True
    Else
        bFontBold = False
        mnuBold.Checked = False
        SaveSetting "ASNIExpress", "Settings", "FontBold", False
    End If
    AdjustScreen
End Sub

Private Sub mnuClear_Click()
    picNorm.Cls
    picBlnk.Cls
    picNorm.CurrentX = 0
    picNorm.CurrentY = 0
    picBlnk.CurrentX = 0
    picBlnk.CurrentY = 0
    x = 0
    Y = 0
End Sub

Private Sub mnuCopy_Click()
    DoEvents
    Dim b As Picture
    Set b = CaptureWindow(picNorm.hwnd, True, 0, 0, xOffSet * xScreenSize, yOffSet * yScreenSize)
    Clipboard.Clear
    Clipboard.SetData b
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuItalic_Click()
    If bFontItalic = False Then
        bFontItalic = True
        mnuItalic.Checked = True
        SaveSetting "ASNIExpress", "Settings", "FontItalic", True
    Else
        bFontItalic = False
        mnuItalic.Checked = False
        SaveSetting "ASNIExpress", "Settings", "FontItalic", False
    End If
    AdjustScreen
End Sub

Private Sub mnuOpen_Click()
On Error GoTo Err_Handler
Dim t As String
    CommonDialog1.ShowOpen
    x = 0
    Y = 0
    picNorm.Cls
    picNorm.CurrentX = 0
    picNorm.CurrentY = 0
    picBlnk.Cls
    picBlnk.CurrentX = 0
    picBlnk.CurrentY = 0
    SaveX = 0
    SaveY = 0
    f = ""
    Open CommonDialog1.FileName For Input As #1
    Do While Not EOF(1)
        Line Input #1, t
        f = f & t & vbCr
    Loop
    Close #1
    f = Left(f, Len(f) - 1)
    f = Replace(f, "[?7h", "")
    Display f
    
Exit Sub
Err_Handler:
    If Err.Number <> 32755 Then
        MsgBox Err.Description, vbCritical, "Error opening file"
    End If
End Sub

Private Sub mnuPrint_Click()
On Error GoTo Err_Handler
    
    Dim b As Picture
    Set b = CaptureWindow(picNorm.hwnd, True, 0, 0, xOffSet * xScreenSize, yOffSet * yScreenSize)
    CommonDialog1.Flags = cdlPDSelection
    CommonDialog1.ShowPrinter
    Printer.PaintPicture b, 0, 0
    Printer.EndDoc
Exit Sub
Err_Handler:
    If Err.Number <> 32755 Then
        MsgBox Err.Description, vbCritical, "Error opening file"
    End If
End Sub

Private Sub mnuRate_Click()
    frmBlinkRate.Show 1
End Sub

Private Sub mnuRedraw_Click()
    picNorm.Cls
    picBlnk.Cls
    picNorm.CurrentX = 0
    picNorm.CurrentY = 0
    picBlnk.CurrentX = 0
    picBlnk.CurrentY = 0
    x = 0
    Y = 0
    Display f
End Sub

Private Sub mnuScreen_Click()
    Dim tx, ty
    tx = xScreenSize
    ty = yScreenSize
    frmScreenSize.Show 1
    If tx <> xScreenSize Or ty <> yScreenSize Then
        ScrollTop = 1
        ScrollBottom = yScreenSize
        AdjustScreen
    End If
End Sub

Private Sub mnuSpeed_Click(Index As Integer)
    Select Case Index
        Case 0
            DisplayRate = 300
        Case 1
            DisplayRate = 1200
        Case 2
            DisplayRate = 2400
        Case 3
            DisplayRate = 9600
        Case 4
            DisplayRate = 14400
        Case 5
            DisplayRate = 28800
        Case 6
            DisplayRate = 33600
        Case 7
            DisplayRate = 56700
        Case 8
            frmDisplayRate.Show 1
        Case 9
            DisplayRate = 0
    End Select
    For i = 0 To 9
        mnuSpeed(i).Checked = False
    Next i
    Select Case DisplayRate
        Case 300
            mnuSpeed(0).Checked = True
        Case 1200
            mnuSpeed(1).Checked = True
        Case 2400
            mnuSpeed(2).Checked = True
        Case 9600
            mnuSpeed(3).Checked = True
        Case 14400
            mnuSpeed(4).Checked = True
        Case 28800
            mnuSpeed(5).Checked = True
        Case 33600
            mnuSpeed(6).Checked = True
        Case 56700
            mnuSpeed(7).Checked = True
        Case 0
            mnuSpeed(9).Checked = True
        Case Else
            mnuSpeed(8).Checked = True
    End Select
    SaveSetting "ASNIExpress", "Settings", "DisplayRate", DisplayRate
    mnuRedraw_Click
End Sub

Private Sub mnuVerySmall_Click()
    mnuVerySmall.Checked = True
    mnuSmall.Checked = False
    mnuNormal.Checked = False
    mnuLarge.Checked = False
    mnuVeryLarge.Checked = False
    SaveSetting "ASNIExpress", "Settings", "TextSize", 1
    picNorm.FontSize = 5
    picBlnk.FontSize = 5
    xOffSet = 4
    yOffSet = 6
    AdjustScreen
End Sub

Private Sub mnuSmall_Click()
    mnuVerySmall.Checked = False
    mnuSmall.Checked = True
    mnuNormal.Checked = False
    mnuLarge.Checked = False
    mnuVeryLarge.Checked = False
    SaveSetting "ASNIExpress", "Settings", "TextSize", 2
    picNorm.FontSize = 6
    picBlnk.FontSize = 6
    xOffSet = 6
    yOffSet = 8
    AdjustScreen
End Sub

Private Sub mnuNormal_Click()
    mnuVerySmall.Checked = False
    mnuSmall.Checked = False
    mnuNormal.Checked = True
    mnuLarge.Checked = False
    mnuVeryLarge.Checked = False
    SaveSetting "ASNIExpress", "Settings", "TextSize", 3
    picNorm.FontSize = 9
    picBlnk.FontSize = 9
    xOffSet = 8
    yOffSet = 12
    AdjustScreen
End Sub

Private Sub mnuLarge_Click()
    mnuVerySmall.Checked = False
    mnuSmall.Checked = False
    mnuNormal.Checked = False
    mnuLarge.Checked = True
    mnuVeryLarge.Checked = False
    SaveSetting "ASNIExpress", "Settings", "TextSize", 4
    picNorm.FontSize = 14
    picBlnk.FontSize = 14
    xOffSet = 10
    yOffSet = 18
    AdjustScreen
End Sub

Private Sub mnuVeryLarge_Click()
    mnuVerySmall.Checked = False
    mnuSmall.Checked = False
    mnuNormal.Checked = False
    mnuLarge.Checked = False
    mnuVeryLarge.Checked = True
    SaveSetting "ASNIExpress", "Settings", "TextSize", 5
    picNorm.FontSize = 12
    picBlnk.FontSize = 12
    xOffSet = 12
    yOffSet = 16
    AdjustScreen
End Sub

Public Sub AdjustScreen()
On Error Resume Next
    picNorm.FontBold = bFontBold
    picNorm.FontItalic = bFontItalic
    picNorm.Width = (xOffSet * xScreenSize) * 15
    picNorm.Height = (yOffSet * yScreenSize) * 15
    
    Me.Width = picNorm.Width + 120
    Me.Height = picNorm.Height + 690
    
    picBlnk.FontBold = bFontBold
    picBlnk.FontItalic = bFontItalic
    picBlnk.Width = (xOffSet * xScreenSize) * 15
    picBlnk.Height = (yOffSet * yScreenSize) * 15
    
    x = 0
    Y = 0
    picNorm.Cls
    picNorm.CurrentX = 0
    picNorm.CurrentY = 0
    picBlnk.Cls
    picBlnk.CurrentX = 0
    picBlnk.CurrentY = 0
    SaveX = 0
    SaveY = 0
    
    Display f
End Sub

Private Sub tmrBlink_Timer()
    If picNorm.Visible = True Then
        picNorm.Visible = False
        picBlnk.Visible = True
    Else
        picNorm.Visible = True
        picBlnk.Visible = False
    End If
End Sub

Public Sub scrollup(numlines As Integer, Top As Integer, bot As Integer)
    If numlines >= bot - Top + 1 Then
        ' Just clear from top to bottom
        picNorm.Line (0, yOffSet * (Top - 1))-(xOffSet * xScreenSize, yOffSet * (bot)), picNorm.FillColor, BF
        picBlnk.Line (0, yOffSet * (Top - 1))-(xOffSet * xScreenSize, yOffSet * (bot)), picBlnk.FillColor, BF
    Else
        ' Copy bot-top+1 - numlines lines from (top+numlines) to (top)
        BitBlt picNorm.hDC, 0, (Top - 1) * yOffSet, xOffSet * xScreenSize, _
          yOffSet * (bot - Top + 1 - numlines), picNorm.hDC, 0, yOffSet * (Top + numlines - 1), vbSrcCopy
        BitBlt picBlnk.hDC, 0, (Top - 1) * yOffSet, xOffSet * xScreenSize, _
          yOffSet * (bot - Top + 1 - numlines), picBlnk.hDC, 0, yOffSet * (Top + numlines - 1), vbSrcCopy
        ' Erase lines bot-numlines to bot
        picNorm.Line (0, yOffSet * (bot - numlines))-(xOffSet * xScreenSize, yOffSet * (bot)), picNorm.FillColor, BF
        picBlnk.Line (0, yOffSet * (bot - numlines))-(xOffSet * xScreenSize, yOffSet * (bot)), picBlnk.FillColor, BF
    End If
End Sub

Public Sub scrolldown(numlines As Integer, Top As Integer, bot As Integer)
    If numlines >= bot - Top + 1 Then
        ' Just clear from top to bottom
        picNorm.Line (0, yOffSet * (Top - 1))-(xOffSet * xScreenSize, yOffSet * (bot)), picNorm.FillColor, BF
        picBlnk.Line (0, yOffSet * (Top - 1))-(xOffSet * xScreenSize, yOffSet * (bot)), picBlnk.FillColor, BF
    Else
        ' Copy bot-top+1 - numlines lines from (top+numlines) to (top)
        BitBlt picNorm.hDC, 0, yOffSet * (Top + numlines - 1), xOffSet * xScreenSize, _
          yOffSet * (bot - Top + 1 - numlines), picNorm.hDC, 0, (Top - 1) * yOffSet, vbSrcCopy
        BitBlt picBlnk.hDC, 0, yOffSet * (Top + numlines - 1), xOffSet * xScreenSize, _
          yOffSet * (bot - Top + 1 - numlines), picBlnk.hDC, 0, (Top - 1) * yOffSet, vbSrcCopy
        ' Erase lines bot-numlines to bot
        picNorm.Line (0, yOffSet * (Top - 1))-(xOffSet * xScreenSize, yOffSet * (Top + numlines - 1)), picNorm.FillColor, BF
        picBlnk.Line (0, yOffSet * (Top - 1))-(xOffSet * xScreenSize, yOffSet * (Top + numlines - 1)), picBlnk.FillColor, BF
    End If
End Sub

Public Sub Scroll()

    BitBlt picNorm.hDC, 0, 0, xOffSet * xScreenSize, (((yOffSet * yScreenSize) - yOffSet)), picNorm.hDC, 0, yOffSet, vbSrcCopy
    BitBlt picBlnk.hDC, 0, 0, xOffSet * xScreenSize, (((yOffSet * yScreenSize) - yOffSet)), picBlnk.hDC, 0, yOffSet, vbSrcCopy
    picNorm.Line (0, (yOffSet * yScreenSize) - yOffSet)-(xOffSet * xScreenSize, yOffSet * yScreenSize), picNorm.FillColor, BF
    picBlnk.Line (0, (yOffSet * yScreenSize) - yOffSet)-(xOffSet * xScreenSize, yOffSet * yScreenSize), picBlnk.FillColor, BF
End Sub

