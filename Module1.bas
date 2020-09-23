Attribute VB_Name = "Module1"
Option Explicit

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetDesktopWindow Lib _
   "user32" () As Long
Public Declare Function GetWindowDC Lib _
   "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" _
   (ByVal hWnd As Long, ByVal hdc As Long) As Long

Public SavePics As Long

Type Part
    X As Integer
    Y As Integer
    drx As Double
    dry As Double
    'B As Long
    Red As Long
    Green As Long
    Blue As Long
End Type

Public QuitGame As Boolean

Public Sub Main()
    Form1.Show
    MainLoop
    Unload Form1
    End
End Sub

Public Sub MainLoop()
    Do
        DoEvents
        If QuitGame Then
            Exit Do
        End If
        Form1.doParts
        If SavePics > -1 Then
            'Save a series of pictures here
            PrintScreen
            SavePicture Form1.buffer.Image, App.Path & "\G" & Format(SavePics, "00") & ".bmp"
            'Stop once you've gotten 32 pictures
            SavePics = SavePics + 1
            If SavePics > 32 Then
                SavePics = -1
                Form1.AutoRedraw = False
            End If
        End If
    Loop
End Sub

Private Sub PrintScreen()

  Dim r As Long
  Dim hWndDesk As Long
  Dim hDCDesk As Long

  Dim LeftDesk As Long
  Dim TopDesk As Long
  Dim WidthDesk As Long
  Dim HeightDesk As Long
   
 'define the screen coordinates (upper
 'corner (0,0) and lower corner (Width, Height)
  LeftDesk = 0
  TopDesk = 0
  WidthDesk = Screen.Width \ Screen.TwipsPerPixelX
  HeightDesk = Screen.Height \ Screen.TwipsPerPixelY
   
 'get the desktop handle and display context
  hWndDesk = GetDesktopWindow()
  hDCDesk = GetWindowDC(hWndDesk)
   
 'copy the desktop to the picture box
  r = BitBlt(Form1.buffer.hdc, 0, 0, _
             WidthDesk, HeightDesk, hDCDesk, _
             LeftDesk, TopDesk, vbSrcCopy)

  r = ReleaseDC(hWndDesk, hDCDesk)

End Sub

