VERSION 5.00
Begin VB.Form frmAniGif 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Animated Gifs & Animated Cursors"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   4110
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TimerAni 
      Left            =   3480
      Top             =   360
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load Animated GIF or CURSOR"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   3855
   End
   Begin VB.PictureBox picAni 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3840
      Left            =   120
      Picture         =   "frmAniGif.frx":0000
      ScaleHeight     =   3840
      ScaleWidth      =   3840
      TabIndex        =   0
      Top             =   840
      Width           =   3840
   End
End
Attribute VB_Name = "frmAniGif"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function RedrawWindow Lib "user32.dll" (ByVal hwnd As Long, ByRef lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function DrawIconEx Lib "user32.dll" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long

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
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

' custom UDT - animated gif frame properties
Private Type AniGifProps
    aPic As StdPicture      ' the image itself
    aInterval As Integer    ' the gif-coded frame interval
    aTop As Long            ' the gif-coded top offset
    aLeft As Long           ' the gif-coded left offset
    aMisc As Long           ' could be remarks or other needed information
End Type
Private Type AniCursorProps
    cFrames As Long
    cInterval As Integer
    cHandle As Long
    cSize As Long
End Type


' collection of gif frames
Private aniFrames() As AniGifProps
Private aniCursor As AniCursorProps

' used to help prevent flicker during animation
Private picBkBuff As StdPicture

' used for RedrawWindow API & also to indicate where
' on the destination DC drawing will take place
Private updateRect As RECT

Private Sub Command1_Click()
' allow user to select an animated gif

TimerAni.Enabled = False
If aniCursor.cHandle Then
    DestroyIcon aniCursor.cHandle
    aniCursor.cHandle = 0
End If

With frmMain.dlgPics
    .Flags = cdlOFNFileMustExist
    .CancelError = True
    .Filter = "GIFs|*.gif|Animated Icon (*.ani)|*.ani"
    .FilterIndex = 0
    .DialogTitle = "Select Animated GIFs or Cursors only"
End With
On Error GoTo ExitRoutine
frmMain.dlgPics.ShowOpen

' loaded a file, but is it an animated gif?
' call function to check & build the animation frames
Dim tmpPic As StdPicture
On Error Resume Next
Set tmpPic = LoadPicture(frmMain.dlgPics.FileName)
If tmpPic.Type = vbPicTypeIcon Then
    Set tmpPic = Nothing
    If LoadAniCursor(frmMain.dlgPics.FileName) = False Then Exit Sub
    TimerAni.Interval = aniCursor.cInterval

Else
    If LoadGif(frmMain.dlgPics.FileName) = False Then Exit Sub
    TimerAni.Interval = aniFrames(0).aInterval

End If
' good to go, lets set up the first frame
TimerAni.Tag = 0
TimerAni.Enabled = True

ExitRoutine:
End Sub

Private Function LoadAniCursor(aniPath As String) As Boolean
' Note that the stdPic object cannot load animated cursors.
' this is just one way of possibly showing an animated cursor
' like an animated gif.

Dim hCursor As Long
Dim nrFrames As Long
Dim cX As Long, cY As Long
    
' we need to load the cursor using APIs 'cause VB can't
hCursor = LoadCursorFromFile(frmMain.dlgPics.FileName)

' this is a short cheat to get the number of frames in the cursor
' Note: a more proper way could be used which should return
' the individual frame intervals. However, we'll use the
' shortcut here & set a static frame interval

' When folloiwng line returns 0, then nrFrames is one too many
Do Until DrawIconEx(picAni.hdc, -100, -100, hCursor, 16, 16, nrFrames, 0, &H3) = 0
    nrFrames = nrFrames + 1
Loop

If nrFrames < 2 Then
    MsgBox "That is not an animated cursor", vbExclamation + vbOKOnly
Else

    With aniCursor
        .cHandle = hCursor
        .cFrames = nrFrames - 1
        .cInterval = 200    ' change to desired interval
        .cSize = 64         ' change to desired dimension
    End With
    
    ' for a flicker free animation, we'll need to create a section
    ' of the source DC that will be drawn over
    CreateBackBuffImg aniCursor.cSize, aniCursor.cSize

    LoadAniCursor = True

End If

End Function


Private Function LoadGif(aniPath As String) As Boolean
' This routine was found on PSC in several posts. Original author unknown, but routine is well
' distributed. I have made modifications to work especially for this application

Rem °°°°°°°°°°°°°°°°°°°°°°°°°
    On Error GoTo ErrHandler
Rem °°°°°°°°°°°°°°°°°°°°°°°°°
    
    Dim fNum As Integer
    Dim MaxX As Long, MaxY As Long
    Dim MaxOffsetX As Long, MaxOffsetY As Long
    Dim imgSize As Long
    Dim imgHeader As String, fileHeader As String
    Dim buf$, picbuf$
    Dim imgCount As Integer
    Dim I&, J&, xOff&, yOff&, TimeWait&
    Dim GifEnd As String
    
    GifEnd = Chr(0) & Chr(33) & Chr(249)            ' flag indicating end of file
    
    fNum = FreeFile
    Open aniPath For Binary Access Read As fNum
        buf = String(LOF(fNum), Chr(0))
        Get #fNum, , buf                                            'Get GIF File into buffer
    Close fNum
    
    I = 1
    J = InStr(1, buf, GifEnd) + 1
    fileHeader = Left(buf, J)
    
    Rem °° Not an Gif File °°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°
    If Left$(fileHeader, 3) <> "GIF" Then
       MsgBox "This file is not an animated GIF file", vbInformation + vbOKOnly
       Exit Function
    End If
    Rem °°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°
        
    
    ' Remove any previous sample images
    Erase aniFrames
    
    I = J + 2
    Do ' Split GIF Files at separate pictures
       ' and load them into Image Array
        J = InStr(I, buf, GifEnd) + 3
        If J > Len(GifEnd) Then
            fNum = FreeFile
            Open "temp.gif" For Binary As fNum
                picbuf = String(Len(fileHeader) + J - I, Chr(0))
                picbuf = fileHeader & Mid(buf, I - 1, J - I)
                Put #fNum, 1, picbuf
                imgHeader = Left(Mid(buf, I - 1, J - I), 16)
            Close fNum
            ReDim Preserve aniFrames(0 To imgCount)
            With aniFrames(imgCount)
                .aInterval = ((Asc(Mid(imgHeader, 4, 1))) + (Asc(Mid(imgHeader, 5, 1)) * 256)) * 10
                Set .aPic = LoadPicture("temp.gif")
                .aLeft = Val(Asc(Mid(imgHeader, 9, 1)) + (Asc(Mid(imgHeader, 10, 1)) * 256&))
                .aTop = Val(Asc(Mid(imgHeader, 9, 1)) + (Asc(Mid(imgHeader, 10, 1)) * 256&))
            End With
            imgCount = imgCount + 1
            I = J
            Kill "temp.gif"
        End If
    Loop Until J = 3
' If there is another Image - Load it
    If I < Len(buf) Then
        fNum = FreeFile
        Open "temp.gif" For Binary As fNum
            picbuf = String(Len(fileHeader) + Len(buf) - I, Chr(0))
            picbuf = fileHeader & Mid(buf, I - 1, Len(buf) - I)
            Put #fNum, 1, picbuf
            imgHeader = Left(Mid(buf, I - 1, Len(buf) - I), 16)
        Close fNum
        ReDim Preserve aniFrames(0 To imgCount)
        With aniFrames(imgCount)
            .aInterval = ((Asc(Mid(imgHeader, 4, 1))) + (Asc(Mid(imgHeader, 5, 1)) * 256)) * 10
            Set .aPic = LoadPicture("temp.gif")
            .aLeft = Val(Asc(Mid(imgHeader, 9, 1)) + (Asc(Mid(imgHeader, 10, 1)) * 256&))
            .aTop = Val(Asc(Mid(imgHeader, 9, 1)) + (Asc(Mid(imgHeader, 10, 1)) * 256&))
        End With
        Kill "temp.gif"
    End If


' let's check & clean up the animated timers if needed
' And also calculate largest image size & frame offsets
For imgCount = 0 To UBound(aniFrames)
    With aniFrames(imgCount)
        
        If .aInterval > 60000 Then  ' no interval > 60000
            .aInterval = 60000
        ElseIf .aInterval = 0 Then  ' no zero intervals (use 100 as default)
            .aInterval = 100
        End If
        
        ' get frame width & add any offset as needed (could be + or -)
        imgSize = ConvertHimetrix2Pixels(.aPic.Width, True) + Abs(.aLeft)
        If imgSize > MaxX Then MaxX = imgSize
        ' get frame height & add any offset as needed (could be + or -)
        imgSize = ConvertHimetrix2Pixels(.aPic.Height, False) + Abs(.aTop)
        If imgSize > MaxY Then MaxY = imgSize
        
        ' check for negative frame offsets & track largest as necessary
        If .aLeft < 0 Then
           If Abs(.aLeft) > MaxOffsetX Then MaxOffsetX = Abs(.aLeft)
        End If
        If .aTop < 0 Then
            If Abs(.aTop) > MaxOffsetY Then MaxOffsetY = Abs(.aTop)
        End If
        
    End With
Next

' in theory only, haven't found any animated gif images with negative offsets
If MaxOffsetX > 0 Or MaxOffsetY > 0 Then
    
    ' we need to shift all offsets by the largest negative offset
    For imgCount = 0 To UBound(aniFrames)
        With aniFrames(imgCount)
            .aLeft = .aLeft + MaxOffsetX
            .aTop = .aTop + MaxOffsetY
        End With
    Next

End If

' for a flicker free animation, we'll need to create a section
' of the source DC that will be drawn over
CreateBackBuffImg MaxX, MaxY

LoadGif = True

Exit Function

Rem °°°°<Femme fatale>°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°
ErrHandler:
    MsgBox "Error No. " & Err.Number & " when reading file", vbCritical
    If fNum Then Close #fNum
    Erase aniFrames
    LoadGif = False
Rem °°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°
End Function

Private Sub CreateBackBuffImg(cX As Long, cY As Long)
' Note that you could modify routine to pass a bitmap or stdPic handle
' This routine is hardcoded with the picture box DC & stdPicture

Dim bkBuffDC As Long, bkBuffBmp As Long, oldBmp As Long
Dim destWidth As Long, destHeight As Long

' get the DC's width & height
destWidth = picAni.Width / Screen.TwipsPerPixelX
destHeight = picAni.Height / Screen.TwipsPerPixelY

' clear the picture box & reset the picBkBuff stdPic object
picAni.Cls
If Not picBkBuff Is Nothing Then Set picBkBuff = Nothing

' just to be safe; clipping should prevent this anyway
' Don't allow larger source dimensions than the source itself
If cX > destWidth Then cX = destWidth
If cY > destHeight Then cY = destHeight

On Error Resume Next    ' want a message box if error occurs here

' we'll calculate the source drawing area so that animated gif
' images are centered in the source
With updateRect
    .Left = (destWidth - cX) \ 2
    .Top = (destHeight - cY) \ 2
    .Right = .Left + cX - 1
    .Bottom = .Top + cY - 1

    ' create a temporary DC
    bkBuffDC = CreateCompatibleDC(Me.hdc)
    
    ' create a blank bitmap to draw on & select into temp DC
    bkBuffBmp = CreateCompatibleBitmap(Me.hdc, cX, cY)
    oldBmp = SelectObject(bkBuffDC, bkBuffBmp)

    ' another tweak... if the hDC is not a VB DC then add zero or
    ' a Type Mismatch or Invalid Procuedure Call error might occur
    picAni.Picture.Render bkBuffDC + 0, 0, 0, cX + 0, cY + 0, _
        ScaleX(.Left, vbPixels, vbHimetric), _
        ScaleY(destHeight - .Top, vbPixels, vbHimetric), _
        ScaleX(cX, vbPixels, vbHimetric), _
        -ScaleY(cY, vbPixels, vbHimetric), ByVal 0&
    
    ' a couple of things about the above function call....
    ' 1) To get the proper vertical segment, you need to subtract the
    '    top edge from the height of the picutre (destHieght - .Top)
    ' 2) The next to last param is still negative
        
End With

If Err Then MsgBox Err.Description  'for testing only

' unselect the temp bmp and then delete the temp DC
SelectObject bkBuffDC, oldBmp
DeleteDC bkBuffDC

' convert bitmap handle to a stdPic object
' the tmp bitmap will be destroyed when stdPic is destroyed
Set picBkBuff = HandleToPicture(bkBuffBmp, True)

End Sub

Private Function ConvertHimetrix2Pixels(vHiMetrix As Long, byWidth As Boolean) As Long
' provided for gee-whiz. ScaleX & ScaleY could be used, but if these routines
' were in a class, then ScaleX & ScaleY could not be used directly
If byWidth Then
    ConvertHimetrix2Pixels = vHiMetrix * 1440 / 2540 / Screen.TwipsPerPixelX
Else
    ConvertHimetrix2Pixels = vHiMetrix * 1440 / 2540 / Screen.TwipsPerPixelY
End If
End Function

Private Function ConvertPixels2Himetrix(vPixels As Long, byWidth As Boolean) As Long
' provided for gee-whiz. ScaleX & ScaleY could be used, but if these routines
' were in a class, then ScaleX & ScaleY could not be used directly
If byWidth Then
    ConvertPixels2Himetrix = vPixels / 1440 * 2540 * Screen.TwipsPerPixelX
Else
    ConvertPixels2Himetrix = vPixels / 1440 * 2540 * Screen.TwipsPerPixelY
End If
End Function

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

Private Sub Form_Unload(Cancel As Integer)
'clean up
TimerAni.Enabled = False
Erase aniFrames
If aniCursor.cHandle Then DestroyIcon aniCursor.cHandle
End Sub

Private Sub TimerAni_Timer()
' function to draw the animated gif using stdPic Render function only
' Tweaked to draw an animated cursor also, but uses API since
'   animated cursors cannot be loaded directly into a stdPic object
'   -- can be done but requires a bunch of other steps & is inefficient & waste of resources

' disable timer for now
TimerAni.Enabled = False

' replace/redraw the drawing section first
With picBkBuff
    ' note the + 0 below. On my machine at least, I will get a type mismatch
    ' error even though the variable type expected is Long & the RECT properties
    ' are Long.  Wierd.
    .Render picAni.hdc, updateRect.Left + 0, updateRect.Top + 0, _
        updateRect.Right - updateRect.Left + 1, updateRect.Bottom - updateRect.Top + 1, _
        0, .Height, .Width, -.Height, ByVal 0&
End With
    
If aniCursor.cHandle Then  ' doing animated cursor; otherwise doing animated gif
   
    ' now draw the next frame
    With aniCursor
        DrawIconEx picAni.hdc, updateRect.Left, updateRect.Top, .cHandle, _
            .cSize, .cSize, Val(TimerAni.Tag), 0, &H3
    End With
   
    If Val(TimerAni.Tag) = aniCursor.cFrames Then TimerAni.Tag = -1
    ' need to update flag for next frame to draw
    TimerAni.Tag = Val(TimerAni.Tag) + 1
    ' set new timer interval
    TimerAni.Interval = Val(aniCursor.cInterval)
    
Else
    ' now draw the next frame
    With aniFrames(Val(TimerAni.Tag))
        .aPic.Render picAni.hdc, updateRect.Left + .aLeft, updateRect.Top + .aTop, _
            ConvertHimetrix2Pixels(.aPic.Width, True), ConvertHimetrix2Pixels(.aPic.Height, False), _
            0, .aPic.Height, .aPic.Width, -.aPic.Height, ByVal 0&
    End With
   
    If Val(TimerAni.Tag) = UBound(aniFrames) Then TimerAni.Tag = -1
    ' need to update flag for next frame to draw
    TimerAni.Tag = Val(TimerAni.Tag) + 1
    ' set new timer interval
    TimerAni.Interval = Val(aniFrames(Val(TimerAni.Tag)).aInterval)

End If

' refresh the drawing area
RedrawWindow picAni.hwnd, updateRect, ByVal 0, 1

' Activate Timer
TimerAni.Enabled = True

End Sub
