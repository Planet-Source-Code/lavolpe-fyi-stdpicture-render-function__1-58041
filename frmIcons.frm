VERSION 5.00
Begin VB.Form frmIcons 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tweaking Icon Drawing Quality"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Select another Icon"
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   3480
      Width           =   1935
   End
   Begin VB.PictureBox picSample 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1815
      Index           =   1
      Left            =   2400
      ScaleHeight     =   1815
      ScaleWidth      =   1935
      TabIndex        =   1
      Top             =   1200
      Width           =   1935
   End
   Begin VB.PictureBox picSample 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1815
      Index           =   0
      Left            =   240
      ScaleHeight     =   1815
      ScaleWidth      =   1935
      TabIndex        =   0
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "The modified routine (right side) is generally better for smaller sizes but worse for larger sizes"
      Height          =   495
      Index           =   2
      Left            =   240
      TabIndex        =   6
      Top             =   720
      Width           =   4095
   End
   Begin VB.Label Label2 
      Caption         =   "Modified Drawing "
      Height          =   255
      Index           =   1
      Left            =   2400
      TabIndex        =   5
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Standard Drawing "
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   $"frmIcons.frx":0000
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   0
      Width           =   4095
   End
End
Attribute VB_Name = "frmIcons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CopyImage Lib "user32.dll" (ByVal hImage As Long, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal uFlags As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" _
    (lpPictDesc As PICTDESC, riid As Any, ByVal fOwn As Long, _
    ipic As IPicture) As Long

Private Const LR_COPYFROMRESOURCE As Long = &H4000

Private Type PICTDESC
    cbSize As Long
    pictType As Long
    hIcon As Long
    hPal As Long
End Type

Private iconSource As StdPicture


Private Sub Command1_Click()
' allow user to select an icon

With frmMain.dlgPics
    .Flags = cdlOFNFileMustExist
    .CancelError = True
    .Filter = "Icons|*.ico"
    .FilterIndex = 0
    .DialogTitle = "Select Icons Only"
End With
On Error GoTo ExitRoutine
frmMain.dlgPics.ShowOpen

Set iconSource = LoadPicture(frmMain.dlgPics.FileName)
picSample(0).Cls
picSample(1).Cls
ShowSample
picSample(0).Refresh
picSample(1).Refresh

ExitRoutine:
End Sub

Private Sub Form_Load()
' use form icon to begin with
Set iconSource = Me.Icon
ShowSample
End Sub

Private Sub ShowSample()

Dim iSize As Integer, Looper As Integer
Dim testPic As StdPicture

Set testPic = iconSource
For Looper = 0 To 1
    picSample(Looper).Cls
    ' create 16x16, 32x32 & 64x64 examples
    For iSize = 1 To 3
        ' on second pass, use the modified icon vs simply stretching existing icon
        If Looper Then GetBetterSize testPic, CInt(iSize * 1.2) * 16
        ' draw the icon
        testPic.Render picSample(Looper).hdc, 0, (CInt(iSize * 1.2) - 1) * 16, CInt(iSize * 1.2) * 16, CInt(iSize * 1.2) * 16, _
            0, testPic.Height, testPic.Width, -testPic.Height, ByVal 0&
    Next
Next
Set testPic = Nothing
End Sub

Private Sub GetBetterSize(newPic As StdPicture, newSize As Long)
' Notes about this API, per MSDN.
' On the personal side, this API tends to return better or equal quality
' on icons at 32x32 or smaller.
' DrawIconEx, Render, etc appear to work better on larger sizes

'CopyImage uses the size in the resource file closest to the desired size.
' This will succeed only if hImage was loaded by LoadIcon or LoadCursor,
' or by LoadImage with the LR_SHARED flag

Dim hImage As Long
Set newPic = Nothing
hImage = CopyImage(iconSource.handle, 1, newSize, newSize, LR_COPYFROMRESOURCE)
If hImage Then
    Set newPic = HandleToPicture(hImage, False)
Else
    Set newPic = iconSource
End If
End Sub

Private Function HandleToPicture(ByVal hHandle As Long, isBitmap As Boolean) As Picture
' Convert an icon/bitmap handle to a Picture object

On Error GoTo ExitRoutine

    Dim pic As PICTDESC
    Dim guid(0 To 3) As Long
    
    ' initialize the PictDesc structure
    pic.cbSize = Len(pic)
    If isBitmap Then pic.pictType = vbPicTypeBitmap Else pic.pictType = vbPicTypeIcon
    pic.hIcon = hHandle
    ' this is the IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
    ' we use an array of Long to initialize it faster
    guid(0) = &H7BF80980
    guid(1) = &H101ABF32
    guid(2) = &HAA00BB8B
    guid(3) = &HAB0C3000
    ' create the picture,
    ' return an object reference right into the function result
    OleCreatePictureIndirect pic, guid(0), True, HandleToPicture

ExitRoutine:
End Function

