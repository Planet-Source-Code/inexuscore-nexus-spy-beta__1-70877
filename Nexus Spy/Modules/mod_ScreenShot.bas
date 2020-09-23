Attribute VB_Name = "Mod_Screenshot"
Option Explicit

'// public win32 api declarations
Public Declare Function OleCreatePictureIndirect Lib "olepro32.dll" _
    (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, _
    IPic As IPicture) As Long

'Public Declare Function BMPToJPG Lib "Converter.dll" _
'                        (ByVal InputFile_Name As String, _
'                        ByVal OutputFile_Name As String, _
'                        ByVal Quality As Long) _
'                        As Integer

Public Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function CreateCompatibleDC Lib "GDI32" (ByVal hdc As Long) As Long

Public Declare Function CreateCompatibleBitmap Lib "GDI32" _
    (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long

Public Declare Function SelectObject Lib "GDI32" (ByVal hdc As Long, _
    ByVal hObject As Long) As Long

Public Declare Function GetDeviceCaps Lib "GDI32" (ByVal hdc As Long, _
    ByVal iCapabilitiy As Long) As Long

Public Declare Function GetSystemPaletteEntries Lib "GDI32" _
    (ByVal hdc As Long, ByVal wStartIndex As Long, _
    ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long

Public Declare Function CreatePalette Lib "GDI32" _
    (lpLogPalette As LOGPALETTE) As Long

Public Declare Function SelectPalette Lib "GDI32" (ByVal hdc As Long, _
    ByVal hPalette As Long, ByVal bForceBackground As Long) As Long

Public Declare Function RealizePalette Lib "GDI32" (ByVal hdc As Long) As Long

Public Declare Function BitBlt Lib "GDI32" (ByVal hDCDest As Long, _
    ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, _
    ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal XSrc As Long, _
    ByVal YSrc As Long, ByVal dwRop As Long) As Long
        
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, _
    ByVal hdc As Long) As Long

Public Declare Function DeleteDC Lib "GDI32" (ByVal hdc As Long) As Long

Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, _
    lpRect As RECT) As Long

Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, _
    lpRect As RECT) As Long

Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long

'// public constants declaration
Public Const RASTERCAPS                 As Long = 38
Public Const RC_PALETTE                 As Long = &H100
Public Const SIZEPALETTE                As Long = 104

'// public data types declaration
Public Type PicBmp
    Size                                    As Long
    Type                                    As Long
    hBmp                                    As Long
    hPal                                    As Long
    Reserved                                As Long
End Type

Public Type GUID
    Data1                                   As Long
    Data2                                   As Integer
    Data3                                   As Integer
    Data4(7)                                As Byte
End Type

Public Type PALETTEENTRY
    peRed                                   As Byte
    peGreen                                 As Byte
    peBlue                                  As Byte
    peFlags                                 As Byte
End Type

Public Type LOGPALETTE
    palVersion                              As Integer
    palNumEntries                           As Integer
    palPalEntry(255)                        As PALETTEENTRY
End Type

Public Type RECT
    Left                                    As Long
    Top                                     As Long
    Right                                   As Long
    Bottom                                  As Long
End Type

'// local variables declarations (used in this module)
Private XOld As Long
Public Function CreateBitmapPicture(ByVal hBmp As Long, ByVal hPal As Long) As Picture
    Rem -> *********************************************************************************************************
    Rem -> CreateBitmapPicture   - Creates a bitmap type Picture object from a bitmap and palette.
    Rem -> hBmp                  - Handle to a bitmap
    Rem -> hPal                  - Handle to a Palette - Can be null if the bitmap doesn't use a palette
    Rem -> Returns               - Returns a Picture object containing the bitmap
    Rem -> *********************************************************************************************************
    Dim r   As Long
    Dim Pic As PicBmp
    Dim IPic          As IPicture
    Dim IID_IDispatch As GUID
    
    With IID_IDispatch
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With
    
    With Pic
        .Size = Len(Pic)
        .Type = vbPicTypeBitmap
        .hBmp = hBmp
        .hPal = hPal
    End With
    
    r = OleCreatePictureIndirect(Pic, IID_IDispatch, 1, IPic)
    Set CreateBitmapPicture = IPic
End Function
Public Function CaptureWindow(ByVal hWndSrc As Long, ByVal bClient As Boolean, ByVal LeftSrc As Long, ByVal TopSrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long) As Picture
    Rem -> *********************************************************************************************************
    Rem -> CaptureWindow                            - Captures any portion of a window.
    Rem -> hWndSrc                                  - Handle to the window to be captured
    Rem -> bClient                                  - If True CaptureWindow captures from the bClient area of the window   - If False CaptureWindow captures from the entire window
    Rem -> LeftSrc, TopSrc, WidthSrc, HeightSrc     - Specify the portion of the window to capture - Dimensions need to be specified in pixels
    Rem -> Returns                                  - Returns a Picture object containing a bitmap of the specified portion of the window that was captured
    Rem -> *********************************************************************************************************
    Dim hDCMemory       As Long
    Dim hBmp            As Long
    Dim hBmpPrev        As Long
    Dim r               As Long
    Dim hDCSrc          As Long
    Dim hPal            As Long
    Dim hPalPrev        As Long
    Dim RasterCapsScrn  As Long
    Dim HasPaletteScrn  As Long
    Dim PaletteSizeScrn As Long
    Dim LogPal          As LOGPALETTE
    
    If bClient Then
        hDCSrc = GetDC(hWndSrc)
    Else
        hDCSrc = GetWindowDC(hWndSrc)
    End If
    
    hDCMemory = CreateCompatibleDC(hDCSrc)
    hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
    hBmpPrev = SelectObject(hDCMemory, hBmp)
    RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS)
    HasPaletteScrn = RasterCapsScrn And RC_PALETTE
    PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE)
    
    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        LogPal.palVersion = &H300
        LogPal.palNumEntries = 256
        r = GetSystemPaletteEntries(hDCSrc, 0, 256, LogPal.palPalEntry(0))
        hPal = CreatePalette(LogPal)
        hPalPrev = SelectPalette(hDCMemory, hPal, 0)
        r = RealizePalette(hDCMemory)
    End If
    
    r = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, LeftSrc, TopSrc, vbSrcCopy)
    hBmp = SelectObject(hDCMemory, hBmpPrev)
    
    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        hPal = SelectPalette(hDCMemory, hPalPrev, 0)
    End If
    
    r = DeleteDC(hDCMemory)
    r = ReleaseDC(hWndSrc, hDCSrc)
    Set CaptureWindow = CreateBitmapPicture(hBmp, hPal)
End Function
Public Function CaptureScreen() As StdPicture
    Dim hWndScreen As Long
    hWndScreen = GetDesktopWindow()
    
    With Screen
        Set CaptureScreen = CaptureWindow(hWndScreen, False, 0, 0, .Width \ .TwipsPerPixelX, .Height \ .TwipsPerPixelY)
    End With
End Function
Public Function CaptureForm(frm As Form) As Picture
    With frm
        Set CaptureForm = CaptureWindow(.hwnd, False, 0, 0, .ScaleX(.Width, vbTwips, vbPixels), .ScaleY(.Height, vbTwips, vbPixels))
    End With
End Function
Public Function CaptureClient(frm As Form) As Picture
    With frm
        Set CaptureClient = CaptureWindow(.hwnd, True, 0, 0, .ScaleX(.ScaleWidth, .ScaleMode, vbPixels), .ScaleY(.ScaleHeight, .ScaleMode, vbPixels))
    End With
End Function
Public Function CaptureActiveWindow() As Picture
    Dim hWndActive As Long
    Dim RectActive As RECT
    Dim blReturn As Long
    
    hWndActive = GetForegroundWindow(): DoEvents
    blReturn = GetWindowRect(hWndActive, RectActive)
    
    With RectActive
        Set CaptureActiveWindow = CaptureWindow(hWndActive, False, 0, 0, .Right - .Left, .Bottom - .Top)
    End With
End Function
Public Sub PrintPictureToFitPage(Prn As Printer, Pic As Picture)
    Rem -> *********************************************************************************************************
    Rem ->  PrintPictureToFitPage - Prints a Picture object as big as possible.
    Rem ->  Prn - Destination Printer object
    Rem ->  Pic - Source Picture object
    Rem -> *********************************************************************************************************
    Dim PicRatio     As Double
    Dim PrnWidth     As Double
    Dim PrnHeight    As Double
    Dim PrnRatio     As Double
    Dim PrnPicWidth  As Double
    Dim PrnPicHeight As Double
    
    Const vbHiMetric As Integer = 8
    
    If Pic.Height >= Pic.Width Then
        Prn.Orientation = vbPRORPortrait   'Taller than wide
    Else
        Prn.Orientation = vbPRORLandscape  'Wider than tall
    End If
    
    PicRatio = Pic.Width / Pic.Height
    
    With Prn
        PrnWidth = .ScaleX(.ScaleWidth, .ScaleMode, vbHiMetric)
        PrnHeight = .ScaleY(.ScaleHeight, .ScaleMode, vbHiMetric)
    End With
    
    PrnRatio = PrnWidth / PrnHeight
    
    If PicRatio >= PrnRatio Then
        PrnPicWidth = Prn.ScaleX(PrnWidth, vbHiMetric, Prn.ScaleMode)
        PrnPicHeight = Prn.ScaleY(PrnWidth / PicRatio, vbHiMetric, Prn.ScaleMode)
    Else
        PrnPicHeight = Prn.ScaleY(PrnHeight, vbHiMetric, Prn.ScaleMode)
        PrnPicWidth = Prn.ScaleX(PrnHeight * PicRatio, vbHiMetric, Prn.ScaleMode)
    End If
    
    Call Prn.PaintPicture(Pic, 0, 0, PrnPicWidth, PrnPicHeight)
End Sub
