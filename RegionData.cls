VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RegionData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Copyright Craig Bonathan, 2004
' Note: You may learn from this code and use it in
'       your own programs, as long as it remains
'       unchanged.

' API calls used to load the region
Private Declare Function SetWindowRgn Lib "user32.dll" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function ExtCreateRegion Lib "gdi32.dll" (ByVal lpXform As Long, ByVal nCount As Long, ByVal lpRgnData As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long


Private TargetWindow As Form, RegionHandle As Long


Public Sub TransformWindow(Window As Form, RegionFile As String, Optional ResizeMode As Long = vbTwips)
    Dim FileNum As Long, RegionBufferSize As Long, RegionBuffer() As Byte
    Dim Width As Long, Height As Long
    
    ' Open the region file and extract the region buffer
    FileNum = FreeFile
    Open RegionFile For Binary Access Read As #FileNum
    RegionBufferSize = LOF(FileNum)
    If RegionBufferSize > 8 Then
        ReDim RegionBuffer(RegionBufferSize - 9)
        Get #FileNum, 1, Width
        Get #FileNum, 5, Height
        Get #FileNum, 9, RegionBuffer
        Close #FileNum
    Else
        Close #FileNum
        Exit Sub
    End If
    
    Set TargetWindow = Window
    
    ' As an option, resize the window to the same size as the region
    If ResizeMode > 0 Then
        TargetWindow.Width = TargetWindow.ScaleX(Width, vbPixels, ResizeMode)
        TargetWindow.Height = TargetWindow.ScaleY(Height, vbPixels, ResizeMode)
    End If
    
    ' Create the region from the region buffer, and apply it to the window
    RegionHandle = ExtCreateRegion(0, RegionBufferSize, VarPtr(RegionBuffer(0)))
    SetWindowRgn TargetWindow.hWnd, RegionHandle, True
End Sub

Private Sub Class_Terminate()
    ' Automatically unload the region when the object is closed
    If RegionHandle <> 0 Then
        SetWindowRgn TargetWindow.hWnd, 0, True
        DeleteObject RegionHandle
    End If
End Sub
