VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form MainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Region Maker"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   5175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton BrowseButton2 
      Caption         =   "..."
      Height          =   255
      Left            =   4680
      TabIndex        =   7
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton BrowseButton1 
      Caption         =   "..."
      Height          =   255
      Left            =   4680
      TabIndex        =   6
      Top             =   120
      Width           =   375
   End
   Begin MSComDlg.CommonDialog FileDialog 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton ViewButton 
      Caption         =   "View Region and Bitmap"
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   840
      Width           =   2415
   End
   Begin VB.CommandButton CreateRegionButton 
      Caption         =   "Make Region from Bitmap"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   2415
   End
   Begin VB.TextBox RegionFileText 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Top             =   480
      Width           =   3495
   End
   Begin VB.TextBox BitmapFileText 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Region File:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Bitmap File:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BrowseButton1_Click()
    FileDialog.FileName = ""
    FileDialog.Filter = "Windows Bitmap (*.bmp)|*.bmp"
    FileDialog.ShowOpen
    If FileDialog.FileName <> "" Then BitmapFileText.Text = FileDialog.FileName
End Sub

Private Sub BrowseButton2_Click()
    FileDialog.FileName = ""
    FileDialog.Filter = "Region File (*.*)|*.*"
    FileDialog.ShowSave
    If FileDialog.FileName <> "" Then RegionFileText.Text = FileDialog.FileName
End Sub

Private Sub CreateRegionButton_Click()
    Dim Colour As Long, Red As Long, Green As Long, Blue As Long
    If BitmapFileText.Text = "" Then Exit Sub
    If RegionFileText.Text = "" Then Exit Sub
    
    If MsgBox("Would you like to set black as the transparent colour?", vbYesNo) = vbYes Then
        Colour = 0
    Else
        Red = CLng(InputBox("Please enter the red value of the transparent colour (0-255):", , "0"))
        Green = CLng(InputBox("Please enter the green value of the transparent colour (0-255):", , "0"))
        Blue = CLng(InputBox("Please enter the blue value of the transparent colour (0-255):", , "0"))
        Colour = RGB(Red, Green, Blue)
    End If
    If MsgBox("For a large image, this may take a while. Continue?", vbYesNo) = vbYes Then
        MainForm.Enabled = False
        CreateRegionFile RegionFileText.Text, BitmapFileText.Text, Colour
        MsgBox ("Region saved")
        MainForm.Enabled = True
    End If
End Sub

Private Sub ViewButton_Click()
    If BitmapFileText.Text = "" Then Exit Sub
    If RegionFileText.Text = "" Then Exit Sub
    
    TestForm.Show
    TestForm.WindowRegion.TransformWindow TestForm, RegionFileText.Text, vbTwips
    Set TestForm.Picture = LoadPicture(BitmapFileText.Text)
End Sub
