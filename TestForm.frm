VERSION 5.00
Begin VB.Form TestForm 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Region Maker Test Window"
   ClientHeight    =   3465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3780
   LinkTopic       =   "Form1"
   ScaleHeight     =   231
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   252
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "TestForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public WindowRegion As New RegionData
Dim OldX As Long, OldY As Long

Private Sub Form_DblClick()
    MainForm.Show
    Unload Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    OldX = X
    OldY = Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Me.Left = Me.Left + Me.ScaleX(X - OldX, vbPixels, vbTwips)
        Me.Top = Me.Top + Me.ScaleY(Y - OldY, vbPixels, vbTwips)
    End If
End Sub
