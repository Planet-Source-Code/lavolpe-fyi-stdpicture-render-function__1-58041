VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "StdPic Render Function"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   3375
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picSample 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   3
      Left            =   2760
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   11
      Top             =   2880
      Width           =   240
   End
   Begin VB.CheckBox chkScale 
      Caption         =   "Show all images to Scale"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   600
      Value           =   1  'Checked
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Other Stuff"
      Height          =   375
      Left            =   2160
      TabIndex        =   9
      Top             =   120
      Width           =   975
   End
   Begin MSComDlg.CommonDialog dlgPics 
      Left            =   2400
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picActual 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      ScaleHeight     =   495
      ScaleWidth      =   3015
      TabIndex        =   7
      Top             =   3480
      Width           =   3015
   End
   Begin VB.PictureBox picSample 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   2
      Left            =   2160
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   5
      Top             =   2640
      Width           =   480
   End
   Begin VB.PictureBox picSample 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   960
      Index           =   1
      Left            =   2160
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   3
      Top             =   1200
      Width           =   960
   End
   Begin VB.PictureBox picSample 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1920
      Index           =   0
      Left            =   120
      ScaleHeight     =   1920
      ScaleWidth      =   1920
      TabIndex        =   1
      Top             =   1200
      Width           =   1920
   End
   Begin VB.CommandButton cmdLoadPic 
      Caption         =   "Load Sample Image"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "16x"
      Height          =   255
      Index           =   4
      Left            =   2760
      TabIndex        =   12
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Actual Size"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   8
      Top             =   3240
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "32 x 32"
      Height          =   255
      Index           =   2
      Left            =   2160
      TabIndex        =   6
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "64 x 64"
      Height          =   255
      Index           =   1
      Left            =   2160
      TabIndex        =   4
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "128 x 128"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1935
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' UPDATED 5 JAN 05

' This project was put together 'cause I wanted to understand the
' Standard Picture object's RENDER function. There is very little
' out there on this function.  Why?

' What I've found...
' It can replace the need of BitBlt, StretchBlt, PaintPicture and
' DrawIcon/DrawIconEx in most cases.  PaintPicture is done from a form,
' picturebox, etc, while Render can be used from any stdPicture object,
' including ones assigned to a picturebox, image control, etc.

' I haven't found an easy way to use RENDER when you only want to transfer
' part of the image to a destination DC.  It is very easy to use when you
' want to tansfer the entire picture.
' ** UPDATED....
' Now found the right formula. It is so easy now that you know the trick
' See the Sub CreateBackBuffImg on frmAniGif for an example.

' The big plus, if you can say that... The RENDER function negates the
' need for other transparency functions when used on transparent GIFs
' and animated GIFs with transparency.

' The examples you'll find here are using RENDER and not, BitBlt, StretchBlt
' or DrawIcon to transfer images.  There is no transparency function in this
' project since it is not needed.

' Exception are present though:
' 1. I used DrawIconEx to draw animated cursor frames. This is simply 'cause
'   the StdPicture cannot be used to load an animated cursor.


' Otherwise, I think you will find that the stdPicture.Render function is
' a good catch-all image blitter and works for all of the following
'   bitmaps
'   icons
'   cursors (except animated ones)
'   gifs (including transparent gifs)
'   jpgs
'   metafiles
'   maybe others; I didn't test any others

' One final note.  You'll see a lot of ScaleX & ScaleY function calls.
' If I were to set the form & picturebox scale propeties to vbPixels
' then those wouldn't be needed at all.  I only left them in for those
' new to VB and are looking for a way to convert scale units.
' Another reason to use ScaleX/ScaleY is that it appears to prevent
' Render from returning a "Type Mismatch" error which it tends to do
' if you don't use it.  See the timer function in frmAniGif for an example.

' also you'll notice that when I tried to Render to a memory DC vs a VB DC,
' I had to use the same trick (memoryDChandle + 0) to get past an
' Invalid Procedure Call error.  See frmAniGif.CreateBackBuffImg for example.



Option Explicit
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" _
    (lpPictDesc As PICTDESC, riid As Any, ByVal fOwn As Long, _
    ipic As IPicture) As Long
Private Declare Function LoadCursorFromFile Lib "user32" Alias "LoadCursorFromFileA" _
    (ByVal lpFileName As String) As Long

Private Type PICTDESC
    cbSize As Long
    pictType As Long
    hIcon As Long
    hPal As Long
End Type


Private sourceSampleImage As StdPicture

Private Sub cmdLoadPic_Click()
With dlgPics
    .Flags = cdlOFNFileMustExist
    .CancelError = True
    .Filter = "All Pictures|*.jpg;*.jpeg;*.gif;*.bmp;*.ico;*.cur;*.bmp;*.wmf|Bitmaps|*bmp|Cursors|*.cur;*.ani|GIFs|*.gif|JPEGs|*.jpg;*.jpeg|Icons|*.ico|Windows Meta Files|*.wmf"
    .DialogTitle = "Select Sample Image"
End With
On Error GoTo ExitRoutine
dlgPics.ShowOpen

On Error Resume Next
Set sourceSampleImage = LoadPicture(dlgPics.FileName)
If Err Then
    MsgBox Err.Description, vbInformation + vbOKOnly, "Failed to Load"
    ' probably an animated cursor or non-pic object
    Err.Clear
    Exit Sub
End If

' following is not needed to load cursors into a stdPicture object
' However, VB fails to carry the color of the cursor (if not black & white)
' This trick converts a cursor handle to a stdPic object via APIs and
'   also carrys over the color of the cursor
'   FYI: it also carrys over the cursor's hot spot
If LCase(Right$(dlgPics.FileTitle, 3)) = "cur" Then
    Dim hCursor As Long
    hCursor = LoadCursorFromFile(dlgPics.FileName)
    If hCursor Then
        Set sourceSampleImage = Nothing
        Set sourceSampleImage = IconToPicture(hCursor)
    End If
End If

ShowSample

ExitRoutine:
If Err Then
    MsgBox Err.Description, vbInformation + vbOKOnly
    Err.Clear
End If
End Sub

Private Sub ResizeSamplePicBox()
' size the "actual size" picture box & the form appropriately

With picActual

    .Move .Left, .Top, ScaleX(sourceSampleImage.Width, vbHimetric, vbTwips), _
        ScaleY(sourceSampleImage.Height, vbHimetric, vbTwips)
    
    If .Width > picSample(1).Width + picSample(1).Left Then
        Me.Width = .Width + .Left + (Me.Width - Me.ScaleWidth) + picActual.Left
    Else
        Me.Width = picSample(1).Width + picSample(1).Left + (Me.Width - Me.ScaleWidth) + picActual.Left
    End If
    
    Me.Height = .Height + .Top + (Me.Height - Me.ScaleHeight)

End With
End Sub

Private Sub ShowSample()
' note that this simple routine does not take into consideration non-square
' images. Therefore Render acts just like StretchBlt or PaintPicture.

' A simple ratio routine is included to show that Render can work with
' non-rectangular sizes also

Dim I As Integer
With sourceSampleImage
    ' show the 3 static size samples (128, 64, 32 & 16 pixels)
    For I = picSample.LBound To picSample.UBound
        
        picSample(I).Cls
        .Render picSample(I).hdc, 0, 0, _
            AdjustSize(picSample(I).Width, picSample(I).Height, True), _
            AdjustSize(picSample(I).Width, picSample(I).Height, False), _
            0, .Height, .Width, -.Height, ByVal 0&
    
        ' note above call uses .Height & -.Height of the stdPicture.
        ' This is needed 'cause stdPicture is stored like DIBs (bottom to top)
        ' If you used 0 & .Height respectively, the image would be rotated 180 degrees.
        ' In addition, the destination coordinates/measurements are always in pixels
        ' & the source coordinates/measurements are in himetrics
        ' These are the only 2 tricks to using stdPic's Render function
    Next
    
    ' show the actual size sample, no scaling needed
    ResizeSamplePicBox
    picActual.Cls
    
    .Render picActual.hdc, 0, 0, _
        ScaleX(.Width, vbHimetric, vbPixels), ScaleY(.Height, vbHimetric, vbPixels), _
        0, .Height, .Width, -.Height, ByVal 0&
    
    Label1(3).Caption = "Actual Size " & CLng(ScaleX(.Width, vbHimetric, vbPixels)) & " x " & CLng(ScaleY(.Height, vbHimetric, vbPixels))

End With
End Sub

Private Function AdjustSize(scaleW As Long, scaleH As Long, byWidth As Boolean) As Long
' simple routine to calculate scale sizes

If chkScale.Value = 1 Then

    Dim ratio1 As Single, ratio2 As Single
    
    With sourceSampleImage
        
        'calcualte width ratio of source image to picBox sample width
        ratio1 = ScaleX(scaleW, vbTwips, vbHimetric) / .Width
        
        'calcualte height ratio of source image to picBox sample height
        ratio2 = ScaleY(scaleH, vbTwips, vbHimetric) / .Height
        
        ' pick the smallest of the two ratios
        If ratio2 < ratio1 Then ratio1 = ratio2
        
        ' now return the adjusted width/height for the proper scale
        If byWidth Then
            AdjustSize = ScaleX(.Width * ratio1, vbHimetric, vbPixels)
        Else
            AdjustSize = ScaleY(.Height * ratio1, vbHimetric, vbPixels)
        End If
        
    End With

Else

    ' not scaling; simply convert the passed measurement from twips to pixels
    
    If byWidth Then
    
        AdjustSize = ScaleX(scaleW, vbTwips, vbPixels)
    
    Else
    
        AdjustSize = ScaleY(scaleH, vbTwips, vbPixels)
        
    End If
    
End If

End Function

Private Sub Command1_Click()
frmAniGif.Show
frmIcons.Show
End Sub

Private Sub Form_Load()
Set sourceSampleImage = Me.Icon
ShowSample
End Sub

Private Function IconToPicture(ByVal hIcon As Long) As Picture
' Convert an icon handle to a Picture object

    Dim pic As PICTDESC
    Dim guid(0 To 3) As Long
    
    ' initialize the PictDesc structure
    pic.cbSize = Len(pic)
    pic.pictType = vbPicTypeIcon
    pic.hIcon = hIcon
    ' this is the IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
    ' we use an array of Long to initialize it faster
    guid(0) = &H7BF80980
    guid(1) = &H101ABF32
    guid(2) = &HAA00BB8B
    guid(3) = &HAB0C3000
    ' create the picture,
    ' return an object reference right into the function result
    OleCreatePictureIndirect pic, guid(0), True, IconToPicture
End Function

Private Sub Form_Unload(Cancel As Integer)
Unload frmAniGif
Unload frmIcons
End Sub

Private Sub chkScale_Click()
ShowSample
End Sub


