Attribute VB_Name = "RegionCreation"
Option Explicit

' Copyright Craig Bonathan, 2004
' Note: You may learn from this code and use it in
'       your own programs, as long as it remains
'       unchanged.

' API calls used to create the region
Private Declare Function CreateRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetRegionData Lib "gdi32.dll" (ByVal hRgn As Long, ByVal dwCount As Long, ByVal lpRgnData As Long) As Long
Private Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function GetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long


Private Sub RemovePixelFromRegion(Handle As Long, X As Long, Y As Long)
    Dim NewHandle As Long
    
    ' Create a region of a single pixel
    NewHandle = CreateRectRgn(X, Y, X + 1, Y + 1)
    
    ' Erase the pixel from the main region
    CombineRgn Handle, Handle, NewHandle, 3
    
    ' Delete the region of the single pixel
    DeleteObject NewHandle
End Sub

Private Function CreateInitialRegion(Width As Long, Height As Long) As Long
    CreateInitialRegion = CreateRectRgn(0, 0, Width, Height)
End Function

Public Sub CreateRegionFile(RegionFile As String, BitmapFile As String, RejectColour As Long)
    Dim TempImage As StdPicture, RegionHandle As Long
    Dim ImageDC As Long, OldImageHandle As Long
    Dim Width As Long, Height As Long, X As Long, Y As Long
    
    Dim RegionBufferSize As Long, RegionBuffer() As Byte, FileNum As Long
    
    ' Load bitmap in to a new device-context in memory (makes reading the image fast)
    Set TempImage = LoadPicture(BitmapFile)
    ImageDC = CreateCompatibleDC(GetDC(0))
    OldImageHandle = SelectObject(ImageDC, TempImage.Handle)
    
    ' Convert the himetric measurements to pixels
    Width = (TempImage.Width * 72) / (Screen.TwipsPerPixelX * 127)
    Height = (TempImage.Height * 72) / (Screen.TwipsPerPixelY * 127)
    
    ' Create a full region (i.e. a box)
    RegionHandle = CreateInitialRegion(Width, Height)
    
    ' For every pixel that matches RejectColour, remove it from the region
    For X = 0 To Width - 1
        For Y = 0 To Height - 1
            If GetPixel(ImageDC, X, Y) = RejectColour Then
                RemovePixelFromRegion RegionHandle, X, Y
            End If
        Next
    Next
    
    ' Gets the required buffer size of the region data
    RegionBufferSize = GetRegionData(RegionHandle, 0, 0)
    
    ' If there is existing region data, then copy it to a byte array, and then in to a file
    If RegionBufferSize > 0 Then
        ReDim RegionBuffer(RegionBufferSize - 1)
        GetRegionData RegionHandle, RegionBufferSize, VarPtr(RegionBuffer(0))
        FileNum = FreeFile
        Open RegionFile For Binary Access Write As #FileNum
        Put #FileNum, 1, Width
        Put #FileNum, 5, Height
        Put #FileNum, 9, RegionBuffer
        Close #FileNum
    End If
    
    ' Delete the region from memory
    DeleteObject RegionHandle
    
    ' Delete the bitmap from memory
    SelectObject ImageDC, OldImageHandle
    DeleteDC ImageDC
End Sub

