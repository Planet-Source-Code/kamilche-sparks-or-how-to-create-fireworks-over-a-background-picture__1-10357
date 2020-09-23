VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Form1"
   ClientHeight    =   5235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9615
   LinkTopic       =   "Form1"
   MousePointer    =   99  'Custom
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   5235
   ScaleWidth      =   9615
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   60
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   75
      Width           =   2760
   End
   Begin VB.PictureBox buffer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4395
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   1
      Top             =   4260
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox bg 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   5505
      Picture         =   "Form1.frx":0046
      ScaleHeight     =   1
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1
      TabIndex        =   0
      Top             =   5100
      Visible         =   0   'False
      Width           =   15
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim LifeSpan As Long      ' Life span of each particle
Dim ParticleColor As Long 'The color of the particle.

Const PartsNum = 400     ' Number of particles

Dim Parts(1 To PartsNum) As Part  ' Array of particles
Dim Mx As Integer   ' Mouse X-Axis position
Dim My As Integer   ' Mouse Y-Axis position
Dim r As Long
Dim G As Long
Dim B As Long
Dim dRed As Long
Dim dGreen As Long
Dim dBlue As Long
Private fps As Long
Private StartTime As Long
Private FrameCount As Long
Private FormHDC As Long
Private BgHDC As Long
Private BufferHDC As Long

Private Sub Form_Load()
    Dim StartPt As Long, X As Long, Y As Long, i As Long
    SavePics = -1
    Randomize Timer
    r = 255
    G = 215
    B = 0
    ParticleColor = RGB(r, G, B)
    dRed = r / 32
    dGreen = G / 32
    dBlue = B / 32
    For i = 1 To UBound(Parts)
        StartPt = Random(1, 32)
        ' Sets when particle will first appear
        Parts(i).Red = StartPt * dRed
        Parts(i).Green = StartPt * dGreen
        Parts(i).Blue = StartPt * dBlue
    Next
    'Set the buffer to the right size.
    X = ScaleX(Form1.Width, vbTwips, vbPixels)
    Y = ScaleY(Form1.Height, vbTwips, vbPixels)
    bg.Move 0, 0, X, Y
    buffer.Move 0, 0, Form1.Width, Form1.Height
    StartTime = Timer
    Form1.Show
    Form1.Picture = LoadPicture(App.Path & "/pic.jpg")
    bg.Picture = Form1.Picture
    FormHDC = Form1.hdc
    BgHDC = bg.hdc
    BufferHDC = buffer.hdc
End Sub

Private Function Random(Lower As Long, Upper As Long) As Long
    Random = Int((Upper * Rnd) + Lower)
End Function

Sub Newpart(ByVal Num As Integer, ByVal X As Integer, ByVal Y As Integer)
' Creates a new particle numer Num at starting point (X,Y)
    ' Set starting X-Axis position
    Parts(Num).X = X
    ' Set starting Y-Axis position
    Parts(Num).Y = Y
    ' Set starting remaining lifetime
    Parts(Num).Red = r
    Parts(Num).Green = G
    Parts(Num).Blue = B
    ' Set particle's movement on the X-Axis
    Parts(Num).drx = ((Rnd * 2) - 1) * 20
    ' Set particle's movement on the Y-Axis
    Parts(Num).dry = ((Rnd * 4) - 1) * 20
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Mx = X
    My = Y
End Sub

Sub doParts()
    Dim t As Long, RetVal As Long
    Dim PixelX As Long, PixelY As Long
    Dim OldColor As Long, i As Long
    'Copy the sparks to the form1
    For i = 1 To UBound(Parts)
        With Parts(i)
            PixelX = ScaleX(.X, vbTwips, vbPixels)
            PixelY = ScaleY(.Y, vbTwips, vbPixels)
            'Erase the old particles
            OldColor = GetPixel(BgHDC, PixelX, PixelY)
            RetVal = SetPixelV(FormHDC, PixelX, PixelY, OldColor)
            'Change particle's location
            .X = .X + .drx
            .Y = .Y + .dry
            PixelX = ScaleX(.X, vbTwips, vbPixels)
            PixelY = ScaleY(.Y, vbTwips, vbPixels)
            ' Decrease particle's remaining lifetime
            .Red = .Red - dRed
            .Green = .Green - dGreen
            .Blue = .Blue - dBlue
            If .Red < 0 Then
                .Red = 0
            End If
            If .Green < 0 Then
                .Green = 0
            End If
            If .Blue < 0 Then
                .Blue = 0
            End If
            If .Red = 0 And .Green = 0 And .Blue = 0 Then
                'Particle is dead
                'Create new particle
                Newpart i, Mx, My
            Else
                'Display the particle
                RetVal = SetPixelV(FormHDC, PixelX, PixelY, RGB(.Red, .Green, .Blue))
            End If
        End With
    Next
    'Form1.Refresh
    FrameCount = FrameCount + 1
    On Error Resume Next
    fps = FrameCount / (Timer - StartTime)
    Text1.Text = fps & " fps"
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim AltDown As Boolean
    AltDown = (Shift And vbAltMask) > 0

    If AltDown Then
        SavePics = 1
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    QuitGame = True
End Sub

