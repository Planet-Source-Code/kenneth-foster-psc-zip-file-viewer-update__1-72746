Attribute VB_Name = "modResizePic"
Option Explicit

' http://www.codenewsgroups.net/group/microsoft.public.vb.general.discussion/topic3286.aspx
' http://edais.mvps.org/

'---------------------------------Start Code------------------
Private Type BITMAP '14 bytes
    bmType                      As Long
    bmWidth                     As Long
    bmHeight                    As Long
    bmWidthBytes                As Long
    bmPlanes                    As Integer
    bmBitsPixel                 As Integer
    bmBits                      As Long
End Type

'Win32 API Function Declarations
Private Declare Function CreateCompatibleDC _
                Lib "gdi32" (ByVal hDC As Long) As Long

Private Declare Function DeleteDC _
                Lib "gdi32" (ByVal hDC As Long) As Long

Private Declare Function DeleteObject _
                Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function GetDC _
                Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function GetDesktopWindow _
                Lib "user32" () As Long

Private Declare Function GetObject _
                Lib "gdi32" _
                Alias "GetObjectA" (ByVal hObject As Long, _
                                    ByVal nCount As Long, _
                                    lpObject As Any) As Long

Private Declare Function SelectObject _
                Lib "gdi32" (ByVal hDC As Long, _
                             ByVal hObject As Long) As Long

Private Declare Function SetStretchBltMode _
                Lib "gdi32" (ByVal hDC As Long, _
                             ByVal nStretchMode As Long) As Long

Private Declare Function StretchBlt _
                Lib "gdi32" (ByVal hDC As Long, _
                             ByVal X As Long, _
                             ByVal Y As Long, _
                             ByVal nWidth As Long, _
                             ByVal nHeight As Long, _
                             ByVal hSrcDC As Long, _
                             ByVal xSrc As Long, _
                             ByVal ySrc As Long, _
                             ByVal nSrcWidth As Long, _
                             ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

'Win32 API Constant Declarations
Private Const STRETCH_HALFTONE = 4

Public Function StretchSourcePictureFromFile(FileName As String, picDest As PictureBox) As StdPicture
    Dim hMemDC      As Long
    Dim hOldBmp     As Long
    Dim hMemWdth    As Long
    Dim hMemHght    As Long
    Dim Bmp         As BITMAP
    Dim nRetVal     As Long
    Dim picSrc      As StdPicture
    Dim OldSM       As ScaleModeConstants
    Dim OldAR       As Boolean
    Dim ScaleFactor As Double
    Dim ShowLeft    As Long
    Dim ShowTop     As Long
    Dim ShowWidth   As Long
    Dim ShowHeight  As Long

    If Len(FileName) = 0 Then
        Beep
        Exit Function
    End If

    'Create the memory DC
    hMemDC = CreateCompatibleDC(GetDC(GetDesktopWindow()))
    'Load the picture
    Set picSrc = LoadPicture(FileName)
    'Assign the picture to the memory DC
    hOldBmp = SelectObject(hMemDC, picSrc.Handle)

    'Get the sizes of the picture
    nRetVal = GetObject(picSrc.Handle, Len(Bmp), Bmp)
    hMemWdth = Bmp.bmWidth
    hMemHght = Bmp.bmHeight

    'Make sure there is a picture
    If (hMemWdth > 0) And (hMemHght > 0) Then

        'Stretch the picture to the picturebox
        With picDest
            'Save the PictureBox's ScaleMode and set it to vbPixels
            OldSM = .ScaleMode
            .ScaleMode = vbPixels

            'Get the largest possible scaling factor
            ScaleFactor = Biggest(hMemWdth / .ScaleHeight, hMemHght / _
                    .ScaleWidth)

            'Get the positions and sizes for the destination picture
            ShowWidth = hMemWdth / ScaleFactor
            ShowHeight = hMemHght / ScaleFactor
            ShowLeft = (.ScaleWidth - ShowWidth) / 2
            ShowTop = (.ScaleHeight - ShowHeight) / 2

            'Save the PictureBox's AutoRedraw and set it to True
            OldAR = .AutoRedraw
            .AutoRedraw = True

            .Picture = LoadPicture()
            .Cls
            nRetVal = SetStretchBltMode(.hDC, STRETCH_HALFTONE)
            nRetVal = StretchBlt(.hDC, ShowLeft, ShowTop, ShowWidth, _
                    ShowHeight, hMemDC, 0, 0, hMemWdth, hMemHght, _
                    vbSrcCopy)
            .ScaleWidth = ShowWidth
            .ScaleHeight = ShowHeight
            If (nRetVal = 0) Then
                Debug.Print "StretchBlt() Error Code " & _
                        Err.LastDllError
            End If

            .Refresh
        End With

        'Reset the PictureBox's ScaleMode
        picDest.ScaleMode = OldSM

        'Reset the PictureBox's AutoRedraw
        picDest.AutoRedraw = OldAR
    End If

    'Return the picture object
    Set StretchSourcePictureFromFile = picSrc

    'Clean up the used memory
    Call SelectObject(hMemDC, hOldBmp)
    Call DeleteDC(hMemDC)
    Set picSrc = Nothing
End Function

Public Sub StretchSourcePictureFromPicture(picSrc As StdPicture, _
                                           picDest As PictureBox)
    Dim hMemDC      As Long
    Dim hOldBmp     As Long
    Dim hMemWdth    As Long
    Dim hMemHght    As Long
    Dim Bmp         As BITMAP
    Dim nRetVal     As Long
    Dim OldSM       As ScaleModeConstants
    Dim OldAR       As Boolean
    Dim ScaleFactor As Double
    Dim ShowLeft    As Long
    Dim ShowTop     As Long
    Dim ShowWidth   As Long
    Dim ShowHeight  As Long

    'Make sure we have a valid picture
    If picSrc.Handle = 0 Then
        Beep
        Exit Sub
    End If

    'Create the memory DC
    hMemDC = CreateCompatibleDC(GetDC(GetDesktopWindow()))
    'Assign the picture to the memory DC
    hOldBmp = SelectObject(hMemDC, picSrc.Handle)

    'Get the sizes of the picture
    nRetVal = GetObject(picSrc.Handle, Len(Bmp), Bmp)
    hMemWdth = Bmp.bmWidth
    hMemHght = Bmp.bmHeight

    'Make sure there is a picture
    If (hMemWdth > 0) And (hMemHght > 0) Then

        'Stretch the picture to the picturebox
        With picDest
            'Save the PictureBox's ScaleMode and set it to vbPixels
            OldSM = .ScaleMode
            .ScaleMode = vbPixels

            'Get the largest possible scaling factor
            ScaleFactor = Biggest(hMemWdth / .ScaleWidth, hMemHght / _
                    .ScaleHeight)

            'Get the positions and sizes for the destination picture
            ShowWidth = hMemWdth / ScaleFactor
            ShowHeight = hMemHght / ScaleFactor
            ShowLeft = (.ScaleWidth - ShowWidth) / 2
            ShowTop = (.ScaleHeight - ShowHeight) / 2
           
            'Save the PictureBox's AutoRedraw and set it to True
            OldAR = .AutoRedraw
            .AutoRedraw = True

            .Cls
            nRetVal = SetStretchBltMode(.hDC, STRETCH_HALFTONE)
            nRetVal = StretchBlt(.hDC, ShowLeft, ShowTop, ShowWidth, _
                    ShowHeight, hMemDC, 0, 0, hMemWdth, hMemHght, _
                    vbSrcCopy)

            If (nRetVal = 0) Then
                Debug.Print "StretchBlt() Error Code " & _
                        Err.LastDllError
            End If

            .Refresh
        End With

        'Reset the PictureBox's ScaleMode
        picDest.ScaleMode = OldSM
        
        'Reset the PictureBox's AutoRedraw
        picDest.AutoRedraw = OldAR
    End If

    'Clean up the used memory
    Call SelectObject(hMemDC, hOldBmp)
    Call DeleteDC(hMemDC)
End Sub

Private Function Biggest(Val1 As Double, _
                         Val2 As Double) As Double
    Biggest = IIf(Val1 >= Val2, Val1, Val2)
End Function

'-------------------------------End Code-------------------------------



