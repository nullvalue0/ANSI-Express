Attribute VB_Name = "modCapture"
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
  
  '--------------------------------------------------------------------
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '
      ' Visual Basic 4.0 16/32 Capture Routines
      '
      ' This module contains several routines for capturing windows into a
      ' picture.  All the routines work on both 16 and 32 bit Windows
      ' platforms.
      ' The routines also have palette support.
      '
      ' CreateBitmapPicture - Creates a picture object from a bitmap and
      ' palette.
      ' CaptureWindow - Captures any window given a window handle.
      ' CaptureActiveWindow - Captures the active window on the desktop.
      ' CaptureForm - Captures the entire form.
      ' CaptureClient - Captures the client area of a form.
      ' CaptureScreen - Captures the entire screen.
      ' PrintPictureToFitPage - prints any picture as big as possible on
      ' the page.
      '
      ' NOTES
      '    - No error trapping is included in these routines.
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '
      Option Explicit
      Option Base 0

      Private Type PALETTEENTRY
         peRed As Byte
         peGreen As Byte
         peBlue As Byte
         peFlags As Byte
      End Type

      Private Type LOGPALETTE
         palVersion As Integer
         palNumEntries As Integer
         palPalEntry(255) As PALETTEENTRY  ' Enough for 256 colors.
      End Type

      Private Type GUID
         Data1 As Long
         Data2 As Integer
         Data3 As Integer
         Data4(7) As Byte
      End Type

      #If Win32 Then

         Private Const RASTERCAPS As Long = 38
         Private Const RC_PALETTE As Long = &H100
         Private Const SIZEPALETTE As Long = 104

         Private Type RECT
            Left As Long
            Top As Long
            Right As Long
            Bottom As Long
         End Type

         Private Declare Function CreateCompatibleDC Lib "GDI32" ( _
            ByVal hDC As Long) As Long
         Private Declare Function CreateCompatibleBitmap Lib "GDI32" ( _
            ByVal hDC As Long, ByVal nWidth As Long, _
            ByVal nHeight As Long) As Long
         Private Declare Function GetDeviceCaps Lib "GDI32" ( _
            ByVal hDC As Long, ByVal iCapabilitiy As Long) As Long
         Private Declare Function GetSystemPaletteEntries Lib "GDI32" ( _
            ByVal hDC As Long, ByVal wStartIndex As Long, _
            ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) _
            As Long
         Private Declare Function CreatePalette Lib "GDI32" ( _
            lpLogPalette As LOGPALETTE) As Long
         Private Declare Function SelectObject Lib "GDI32" ( _
            ByVal hDC As Long, ByVal hObject As Long) As Long
         Public Declare Function BitBlt Lib "GDI32" ( _
            ByVal hDCDest As Long, ByVal XDest As Long, _
            ByVal YDest As Long, ByVal nWidth As Long, _
            ByVal nHeight As Long, ByVal hDCSrc As Long, _
            ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) _
            As Long
         Private Declare Function DeleteDC Lib "GDI32" ( _
            ByVal hDC As Long) As Long
         Private Declare Function GetForegroundWindow Lib "user32" () _
            As Long
         Private Declare Function SelectPalette Lib "GDI32" ( _
            ByVal hDC As Long, ByVal hPalette As Long, _
            ByVal bForceBackground As Long) As Long
         Private Declare Function RealizePalette Lib "GDI32" ( _
            ByVal hDC As Long) As Long
         Private Declare Function GetWindowDC Lib "user32" ( _
            ByVal hwnd As Long) As Long
         Private Declare Function GetDC Lib "user32" ( _
            ByVal hwnd As Long) As Long
         Private Declare Function GetWindowRect Lib "user32" ( _
            ByVal hwnd As Long, lpRect As RECT) As Long
         Private Declare Function ReleaseDC Lib "user32" ( _
            ByVal hwnd As Long, ByVal hDC As Long) As Long
         Private Declare Function GetDesktopWindow Lib "user32" () As Long

         Private Type PicBmp
            Size As Long
            Type As Long
            hBmp As Long
            hPal As Long
            Reserved As Long
         End Type

         Private Declare Function OleCreatePictureIndirect _
            Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, _
            ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long

      #ElseIf Win16 Then

         Private Const RASTERCAPS As Integer = 38
         Private Const RC_PALETTE As Integer = &H100
         Private Const SIZEPALETTE As Integer = 104

         Private Type RECT
            Left As Integer
            Top As Integer
            Right As Integer
            Bottom As Integer
         End Type

         Private Declare Function CreateCompatibleDC Lib "GDI" ( _
            ByVal hDC As Integer) As Integer
         Private Declare Function CreateCompatibleBitmap Lib "GDI" ( _
            ByVal hDC As Integer, ByVal nWidth As Integer, _
            ByVal nHeight As Integer) As Integer
         Private Declare Function GetDeviceCaps Lib "GDI" ( _
            ByVal hDC As Integer, ByVal iCapabilitiy As Integer) As Integer
         Private Declare Function GetSystemPaletteEntries Lib "GDI" ( _
            ByVal hDC As Integer, ByVal wStartIndex As Integer, _
            ByVal wNumEntries As Integer, _
            lpPaletteEntries As PALETTEENTRY) As Integer
         Private Declare Function CreatePalette Lib "GDI" ( _
            lpLogPalette As LOGPALETTE) As Integer
         Private Declare Function SelectObject Lib "GDI" ( _
            ByVal hDC As Integer, ByVal hObject As Integer) As Integer
         Private Declare Function BitBlt Lib "GDI" ( _
            ByVal hDCDest As Integer, ByVal XDest As Integer, _
            ByVal YDest As Integer, ByVal nWidth As Integer, _
            ByVal nHeight As Integer, ByVal hDCSrc As Integer, _
            ByVal XSrc As Integer, ByVal YSrc As Integer, _
            ByVal dwRop As Long) As Integer
         Private Declare Function DeleteDC Lib "GDI" ( _
            ByVal hDC As Integer) As Integer
         Private Declare Function GetForegroundWindow Lib "USER" _
            Alias "GetActiveWindow" () As Integer
         Private Declare Function SelectPalette Lib "USER" ( _
            ByVal hDC As Integer, ByVal hPalette As Integer, ByVal _
            bForceBackground As Integer) As Integer
         Private Declare Function RealizePalette Lib "USER" ( _
            ByVal hDC As Integer) As Integer
         Private Declare Function GetWindowDC Lib "USER" ( _
            ByVal hwnd As Integer) As Integer
         Private Declare Function GetDC Lib "USER" ( _
            ByVal hwnd As Integer) As Integer
         Private Declare Function GetWindowRect Lib "USER" ( _
            ByVal hwnd As Integer, lpRect As RECT) As Integer
         Private Declare Function ReleaseDC Lib "USER" ( _
            ByVal hwnd As Integer, ByVal hDC As Integer) As Integer
         Private Declare Function GetDesktopWindow Lib "USER" () As Integer

         Private Type PicBmp
            Size As Integer
            Type As Integer
            hBmp As Integer
            hPal As Integer
            Reserved As Integer
         End Type

         Private Declare Function OleCreatePictureIndirect _
            Lib "oc25.dll" (PictDesc As PicBmp, RefIID As GUID, _
            ByVal fPictureOwnsHandle As Integer, IPic As IPicture) _
            As Integer

      #End If

      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '
      ' CreateBitmapPicture
      '    - Creates a bitmap type Picture object from a bitmap and
      '      palette.
      '
      ' hBmp
      '    - Handle to a bitmap.
      '
      ' hPal
      '    - Handle to a Palette.
      '    - Can be null if the bitmap doesn't use a palette.
      '
      ' Returns
      '    - Returns a Picture object containing the bitmap.
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '
      #If Win32 Then
         Public Function CreateBitmapPicture(ByVal hBmp As Long, _
            ByVal hPal As Long) As Picture

            Dim r As Long
      #ElseIf Win16 Then
         Public Function CreateBitmapPicture(ByVal hBmp As Integer, _
            ByVal hPal As Integer) As Picture

            Dim r As Integer
      #End If
         Dim Pic As PicBmp
         ' IPicture requires a reference to "Standard OLE Types."
         Dim IPic As IPicture
         Dim IID_IDispatch As GUID

         ' Fill in with IDispatch Interface ID.
         With IID_IDispatch
            .Data1 = &H20400
            .Data4(0) = &HC0
            .Data4(7) = &H46
         End With

         ' Fill Pic with necessary parts.
         With Pic
            .Size = Len(Pic)          ' Length of structure.
            .Type = vbPicTypeBitmap   ' Type of Picture (bitmap).
            .hBmp = hBmp              ' Handle to bitmap.
            .hPal = hPal              ' Handle to palette (may be null).
         End With

         ' Create Picture object.
         r = OleCreatePictureIndirect(Pic, IID_IDispatch, 1, IPic)

         ' Return the new Picture object.
         Set CreateBitmapPicture = IPic
      End Function

      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '
      ' CaptureWindow
      '    - Captures any portion of a window.
      '
      ' hWndSrc
      '    - Handle to the window to be captured.
      '
      ' Client
      '    - If True CaptureWindow captures from the client area of the
      '      window.
      '    - If False CaptureWindow captures from the entire window.
      '
      ' LeftSrc, TopSrc, WidthSrc, HeightSrc
      '    - Specify the portion of the window to capture.
      '    - Dimensions need to be specified in pixels.
      '
      ' Returns
      '    - Returns a Picture object containing a bitmap of the specified
      '      portion of the window that was captured.
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      ''''''
      '
      #If Win32 Then
         Public Function CaptureWindow(ByVal hWndSrc As Long, _
            ByVal Client As Boolean, ByVal LeftSrc As Long, _
            ByVal TopSrc As Long, ByVal WidthSrc As Long, _
            ByVal HeightSrc As Long) As Picture

            Dim hDCMemory As Long
            Dim hBmp As Long
            Dim hBmpPrev As Long
            Dim r As Long
            Dim hDCSrc As Long
            Dim hPal As Long
            Dim hPalPrev As Long
            Dim RasterCapsScrn As Long
            Dim HasPaletteScrn As Long
            Dim PaletteSizeScrn As Long
      #ElseIf Win16 Then
         Public Function CaptureWindow(ByVal hWndSrc As Integer, _
            ByVal Client As Boolean, ByVal LeftSrc As Integer, _
            ByVal TopSrc As Integer, ByVal WidthSrc As Long, _
            ByVal HeightSrc As Long) As Picture

            Dim hDCMemory As Integer
            Dim hBmp As Integer
            Dim hBmpPrev As Integer
            Dim r As Integer
            Dim hDCSrc As Integer
            Dim hPal As Integer
            Dim hPalPrev As Integer
            Dim RasterCapsScrn As Integer
            Dim HasPaletteScrn As Integer
            Dim PaletteSizeScrn As Integer
      #End If
         Dim LogPal As LOGPALETTE

         ' Depending on the value of Client get the proper device context.
         If Client Then
            hDCSrc = GetDC(hWndSrc) ' Get device context for client area.
         Else
            hDCSrc = GetWindowDC(hWndSrc) ' Get device context for entire
                                          ' window.
         End If

         ' Create a memory device context for the copy process.
         hDCMemory = CreateCompatibleDC(hDCSrc)
         ' Create a bitmap and place it in the memory DC.
         hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
         hBmpPrev = SelectObject(hDCMemory, hBmp)

         ' Get screen properties.
         RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS) ' Raster
                                                            ' capabilities.
         HasPaletteScrn = RasterCapsScrn And RC_PALETTE       ' Palette
                                                              ' support.
         PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE) ' Size of
                                                              ' palette.

         ' If the screen has a palette make a copy and realize it.
         If HasPaletteScrn And (PaletteSizeScrn = 256) Then
            ' Create a copy of the system palette.
            LogPal.palVersion = &H300
            LogPal.palNumEntries = 256
            r = GetSystemPaletteEntries(hDCSrc, 0, 256, _
                LogPal.palPalEntry(0))
            hPal = CreatePalette(LogPal)
            ' Select the new palette into the memory DC and realize it.
            hPalPrev = SelectPalette(hDCMemory, hPal, 0)
            r = RealizePalette(hDCMemory)
         End If

         ' Copy the on-screen image into the memory DC.
         r = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, _
            LeftSrc, TopSrc, vbSrcCopy)

      ' Remove the new copy of the  on-screen image.
         hBmp = SelectObject(hDCMemory, hBmpPrev)
         
         ' If the screen has a palette get back the palette that was
         ' selected in previously.
         If HasPaletteScrn And (PaletteSizeScrn = 256) Then
            hPal = SelectPalette(hDCMemory, hPalPrev, 0)
         End If

         ' Release the device context resources back to the system.
         r = DeleteDC(hDCMemory)
         r = ReleaseDC(hWndSrc, hDCSrc)

         ' Call CreateBitmapPicture to create a picture object from the
         ' bitmap and palette handles. Then return the resulting picture
         ' object.
         Set CaptureWindow = CreateBitmapPicture(hBmp, hPal)
      End Function

      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '
      ' CaptureScreen
      '    - Captures the entire screen.
      '
      ' Returns
      '    - Returns a Picture object containing a bitmap of the screen.
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '
      Public Function CaptureScreen() As Picture
         #If Win32 Then
            Dim hWndScreen As Long
         #ElseIf Win16 Then
            Dim hWndScreen As Integer
         #End If

         ' Get a handle to the desktop window.
         hWndScreen = GetDesktopWindow()

         ' Call CaptureWindow to capture the entire desktop give the handle
         ' and return the resulting Picture object.

         Set CaptureScreen = CaptureWindow(hWndScreen, False, 0, 0, _
            Screen.Width \ Screen.TwipsPerPixelX, _
            Screen.Height \ Screen.TwipsPerPixelY)
      End Function

      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '
      ' CaptureForm
      '    - Captures an entire form including title bar and border.
      '
      ' frmSrc
      '    - The Form object to capture.
      '
      ' Returns
      '    - Returns a Picture object containing a bitmap of the entire
      '      form.
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '
      Public Function CaptureForm(frmSrc As Form) As Picture
         ' Call CaptureWindow to capture the entire form given its window
         ' handle and then return the resulting Picture object.
         Set CaptureForm = CaptureWindow(frmSrc.hwnd, False, 0, 0, _
            frmSrc.ScaleX(frmSrc.Width, vbTwips, vbPixels), _
            frmSrc.ScaleY(frmSrc.Height, vbTwips, vbPixels))
      End Function

      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '
      ' CaptureClient
      '    - Captures the client area of a form.
      '
      ' frmSrc
      '    - The Form object to capture.
      '
      ' Returns
      '    - Returns a Picture object containing a bitmap of the form's
      '      client area.
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '
      Public Function CaptureClient(frmSrc As Form) As Picture
         ' Call CaptureWindow to capture the client area of the form given
         ' its window handle and return the resulting Picture object.
         Set CaptureClient = CaptureWindow(frmSrc.hwnd, True, 0, 0, _
            frmSrc.ScaleX(frmSrc.ScaleWidth, frmSrc.ScaleMode, vbPixels), _
            frmSrc.ScaleY(frmSrc.ScaleHeight, frmSrc.ScaleMode, vbPixels))
      End Function

      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '
      ' CaptureActiveWindow
      '    - Captures the currently active window on the screen.
      '
      ' Returns
      '    - Returns a Picture object containing a bitmap of the active
      '      window.
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '
      Public Function CaptureActiveWindow() As Picture
         #If Win32 Then
            Dim hWndActive As Long
            Dim r As Long
         #ElseIf Win16 Then
            Dim hWndActive As Integer
            Dim r As Integer
         #End If
         Dim RectActive As RECT

         ' Get a handle to the active/foreground window.
         hWndActive = GetForegroundWindow()

         ' Get the dimensions of the window.
         r = GetWindowRect(hWndActive, RectActive)

         ' Call CaptureWindow to capture the active window given its
         ' handle and return the Resulting Picture object.
      Set CaptureActiveWindow = CaptureWindow(hWndActive, False, 0, 0, _
            RectActive.Right - RectActive.Left, _
            RectActive.Bottom - RectActive.Top)
      End Function

      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '
      ' PrintPictureToFitPage
      '    - Prints a Picture object as big as possible.
      '
      ' Prn
      '    - Destination Printer object.
      '
      ' Pic
      '    - Source Picture object.
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '
      Public Sub PrintPictureToFitPage(Prn As Printer, Pic As Picture)
         Const vbHiMetric As Integer = 8
         Dim PicRatio As Double
         Dim PrnWidth As Double
         Dim PrnHeight As Double
         Dim PrnRatio As Double
         Dim PrnPicWidth As Double
         Dim PrnPicHeight As Double

         ' Determine if picture should be printed in landscape or portrait
         ' and set the orientation.
         If Pic.Height >= Pic.Width Then
            Prn.Orientation = vbPRORPortrait   ' Taller than wide.
         Else
            Prn.Orientation = vbPRORLandscape  ' Wider than tall.
         End If

         ' Calculate device independent Width-to-Height ratio for picture.
         PicRatio = Pic.Width / Pic.Height

         ' Calculate the dimentions of the printable area in HiMetric.
         PrnWidth = Prn.ScaleX(Prn.ScaleWidth, Prn.ScaleMode, vbHiMetric)
         PrnHeight = Prn.ScaleY(Prn.ScaleHeight, Prn.ScaleMode, vbHiMetric)
         ' Calculate device independent Width to Height ratio for printer.
         PrnRatio = PrnWidth / PrnHeight

         ' Scale the output to the printable area.
         If PicRatio >= PrnRatio Then
            ' Scale picture to fit full width of printable area.
            PrnPicWidth = Prn.ScaleX(PrnWidth, vbHiMetric, Prn.ScaleMode)
            PrnPicHeight = Prn.ScaleY(PrnWidth / PicRatio, vbHiMetric, _
               Prn.ScaleMode)
         Else
            ' Scale picture to fit full height of printable area.
            PrnPicHeight = Prn.ScaleY(PrnHeight, vbHiMetric, Prn.ScaleMode)
            PrnPicWidth = Prn.ScaleX(PrnHeight * PicRatio, vbHiMetric, _
               Prn.ScaleMode)
         End If

         ' Print the picture using the PaintPicture method.
         Prn.PaintPicture Pic, 0, 0, PrnPicWidth, PrnPicHeight
      End Sub
      '--------------------------------------------------------------------
                    

