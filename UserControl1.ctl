VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.UserControl Marquee 
   BackColor       =   &H80000008&
   ClientHeight    =   450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3405
   ScaleHeight     =   450
   ScaleWidth      =   3405
   ToolboxBitmap   =   "UserControl1.ctx":0000
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   0
      Top             =   0
   End
   Begin PicClip.PictureClip PicClip1 
      Left            =   360
      Top             =   720
      _ExtentX        =   52599
      _ExtentY        =   820
      _Version        =   393216
      Cols            =   70
      Picture         =   "UserControl1.ctx":0312
   End
   Begin PicClip.PictureClip PicClip2 
      Left            =   360
      Top             =   1200
      _ExtentX        =   52599
      _ExtentY        =   820
      _Version        =   393216
      Cols            =   70
      Picture         =   "UserControl1.ctx":F820
   End
   Begin PicClip.PictureClip PicClip3 
      Left            =   360
      Top             =   1680
      _ExtentX        =   52599
      _ExtentY        =   820
      _Version        =   393216
      Cols            =   70
      Picture         =   "UserControl1.ctx":1ED2E
   End
   Begin PicClip.PictureClip PicClip4 
      Left            =   360
      Top             =   2160
      _ExtentX        =   52599
      _ExtentY        =   820
      _Version        =   393216
      Cols            =   70
      Picture         =   "UserControl1.ctx":2E23C
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   0
      Left            =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "Marquee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const YellowDigit = 4
Private Const RedDigit = 3
Private Const GreenDigit = 2
Private Const BlueDigit = 1
Private Const sCaption = "PANOS MARQUEE"
Private Const DEF_Interval = 1000
Dim i, Chars As Integer, DEF_Width, dig As Integer, Cptn As String
Dim Intrvl As Integer, Looped As Boolean, temp As String
Dim s, k, d
Enum DColors
BlueDigits = 1
GreenDigits = 2
RedDigits = 3
YellowDigits = 4
End Enum
Private Sub Timer1_Timer()
If Looped Then
d = d + 1
If d > Len(temp) Then d = 1
k = Mid(temp, d, Chars)
For i = 1 To Chars
s = Mid(k, i, 1)
        Select Case dig
        Case 1
        Image1(i).Picture = PicClip1.GraphicCell(GetLetter(s))
        Case 2
        Image1(i).Picture = PicClip2.GraphicCell(GetLetter(s))
        Case 3
        Image1(i).Picture = PicClip3.GraphicCell(GetLetter(s))
        Case 4
        Image1(i).Picture = PicClip4.GraphicCell(GetLetter(s))
        End Select
Next
Else
d = d + 1
If d >= Len(temp) Then d = 1
k = Mid(temp, Len(temp) - d, Chars)
For i = 1 To Chars
s = Mid(k, i, 1)
        Select Case dig
        Case 1
        Image1(i).Picture = PicClip1.GraphicCell(GetLetter(s))
        Case 2
        Image1(i).Picture = PicClip2.GraphicCell(GetLetter(s))
        Case 3
        Image1(i).Picture = PicClip3.GraphicCell(GetLetter(s))
        Case 4
        Image1(i).Picture = PicClip4.GraphicCell(GetLetter(s))
        End Select

Next

End If

End Sub

Private Sub UserControl_Initialize()
Image1(0).Picture = PicClip1.GraphicCell(0)
Image1(0).Top = 0
DEF_Width = Image1(0).Width
Image1(0).Left = 0 - Image1(0).Width
UserControl.Height = Image1(0).Height
UserControl.Width = 0
Chars = 0
dig = 1
ShowDigits
Digitcolor = BlueDigit
Caption = sCaption
Interval = DEF_Interval
LoopFromLeft = True
UserControl.Width = DEF_Width * Chars
End Sub

Private Sub UserControl_Resize()
Dim New_Width
On Error Resume Next
If UserControl.Height <> Image1(0).Height Then UserControl.Height = Image1(0).Height
If UserControl.Width > DEF_Width * Chars Then
    New_Width = (UserControl.Width - (DEF_Width * Chars)) \ DEF_Width
    If New_Width = 0 Then New_Width = 1
    For i = Chars + 1 To Chars + New_Width
        Load Image1(i)
        Image1(i).Left = Image1(i - 1).Left + Image1(i - 1).Width
        Select Case dig
        Case 1
        Image1(i).Picture = PicClip1.GraphicCell(46)
        Case 2
        Image1(i).Picture = PicClip2.GraphicCell(46)
        Case 3
        Image1(i).Picture = PicClip3.GraphicCell(46)
        Case 4
        Image1(i).Picture = PicClip4.GraphicCell(46)
        End Select
        Image1(i).Visible = True
    Next
    Chars = Chars + New_Width
    UserControl.Width = DEF_Width * Chars
ElseIf UserControl.Width < DEF_Width * Chars Then
    New_Width = ((DEF_Width * Chars) - UserControl.Width) \ DEF_Width
    If New_Width = 0 Then
        Unload Image1(Chars)
        Chars = Chars - 1
        UserControl.Width = DEF_Width * Chars
    Else
        RemoveChars (Chars - New_Width)
        Chars = Chars - New_Width
        UserControl.Width = DEF_Width * Chars
    End If
End If

Exit Sub

panos:
Exit Sub
End Sub

Private Sub RemoveChars(sChars As Integer)
For i = Chars To sChars Step -1
Unload Image1(i)
Next
End Sub


Private Function GetLetter(ByVal ooo As String) As Integer
Select Case ooo
Case "0"
GetLetter = 0
Case "1"
GetLetter = 1
Case "2"
GetLetter = 2
Case "3"
GetLetter = 3
Case "4"
GetLetter = 4
Case "5"
GetLetter = 5
Case "6"
GetLetter = 6
Case "7"
GetLetter = 7
Case "8"
GetLetter = 8
Case "9"
GetLetter = 9
Case "A"
GetLetter = 10
Case "Á"
GetLetter = 10
Case "B"
GetLetter = 11
Case "Â"
GetLetter = 11
Case "C"
GetLetter = 12
Case "D"
GetLetter = 13
Case "E"
GetLetter = 14
Case "Å"
GetLetter = 14
Case "F"
GetLetter = 15
Case "G"
GetLetter = 16
Case "H"
GetLetter = 17
Case "Ç"
GetLetter = 17
Case "I"
GetLetter = 18
Case "É"
GetLetter = 18
Case "J"
GetLetter = 19
Case "K"
GetLetter = 20
Case "Ê"
GetLetter = 20
Case "L"
GetLetter = 21
Case "M"
GetLetter = 22
Case "Ì"
GetLetter = 22
Case "N"
GetLetter = 23
Case "Í"
GetLetter = 23
Case "O"
GetLetter = 24
Case "Ï"
GetLetter = 24
Case "P"
GetLetter = 25
Case "Ñ"
GetLetter = 25
Case "Q"
GetLetter = 26
Case "R"
GetLetter = 27
Case "S"
GetLetter = 28
Case "T"
GetLetter = 29
Case "Ô"
GetLetter = 29
Case "U"
GetLetter = 30
Case "V"
GetLetter = 45
Case "W"
GetLetter = 31
Case "X"
GetLetter = 32
Case "×"
GetLetter = 32
Case "Y"
GetLetter = 33
Case "Õ"
GetLetter = 33
Case "Z"
GetLetter = 34
Case "Æ"
GetLetter = 34
Case "Ã"
GetLetter = 35
Case "Ä"
GetLetter = 36
Case "È"
GetLetter = 37
Case "Ë"
GetLetter = 38
Case "Î"
GetLetter = 39
Case "Ð"
GetLetter = 40
Case "Ó"
GetLetter = 41
Case "Ö"
GetLetter = 42
Case "Ø"
GetLetter = 43
Case "Ù"
GetLetter = 44
Case " "
GetLetter = 46
Case "."
GetLetter = 47
Case ","
GetLetter = 48
Case ":"
GetLetter = 49
Case ";"
GetLetter = 50
Case "+"
GetLetter = 51
Case "-"
GetLetter = 52
Case "="
GetLetter = 53
Case "$"
GetLetter = 54
Case "!"
GetLetter = 55
Case "#"
GetLetter = 56
Case "&"
GetLetter = 57
Case "("
GetLetter = 58
Case ")"
GetLetter = 59
Case "?"
GetLetter = 60
Case "*"
GetLetter = 61
Case "/"
GetLetter = 62
Case "%"
GetLetter = 63
Case "<"
GetLetter = 64
Case ">"
GetLetter = 65
Case "["
GetLetter = 66
Case "]"
GetLetter = 67
Case "{"
GetLetter = 68
Case "}"
GetLetter = 69
Case Else
GetLetter = 46
End Select
End Function

Public Property Get Digitcolor() As DColors
Attribute Digitcolor.VB_Description = "Sets the Digit Color possible values 1 - 4"
If dig > 4 Then dig = 1
Digitcolor = dig
End Property
Public Property Let Digitcolor(sColor As DColors)
    dig = sColor
    If dig > 4 Or dig < 1 Then
    dig = 1
    MsgBox "Use 1 for Blue Digits, 2 for Green Digits, 3 for Red Digits or 4 for Yellow Digits", vbExclamation, App.Title
    End If
    'DeleteDigits
    RefDigits
    PropertyChanged "DigitColor"

End Property

Private Sub ShowDigits()
For i = 1 To 10
Chars = Chars + 1
Load Image1(i)
Image1(i).Left = Image1(i - 1).Left + Image1(i - 1).Width
Select Case dig
Case 1
Image1(i).Picture = PicClip1.GraphicCell(46)
Case 2
Image1(i).Picture = PicClip2.GraphicCell(46)
Case 3
Image1(i).Picture = PicClip3.GraphicCell(46)
Case 4
Image1(i).Picture = PicClip4.GraphicCell(46)
End Select
Image1(i).Visible = True
Next

End Sub

Private Sub DeleteDigits()
For i = Chars To 1 Step -1
Unload Image1(i)
Next
Chars = 0
End Sub

Private Sub RefDigits()
For i = 1 To Chars
Select Case dig
Case 1
Image1(i).Picture = PicClip1.GraphicCell(46)
Case 2
Image1(i).Picture = PicClip2.GraphicCell(46)
Case 3
Image1(i).Picture = PicClip3.GraphicCell(46)
Case 4
Image1(i).Picture = PicClip4.GraphicCell(46)
End Select
Next

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error GoTo ReadPropErr
    Digitcolor = PropBag.ReadProperty("DigitColor", YellowDigit)
    Caption = PropBag.ReadProperty("Caption", sCaption)
    Interval = PropBag.ReadProperty("Interval", 1000)
    LoopFromLeft = PropBag.ReadProperty("LoopFromLeft", True)
EndReadProp:
    Exit Sub
ReadPropErr:
    'Use default property settings
    Resume EndReadProp
End Sub


'Save control properties
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "DigitColor", Digitcolor, YellowDigit
    PropBag.WriteProperty "Caption", Caption, sCaption
    PropBag.WriteProperty "Interval", Interval, DEF_Interval
    PropBag.WriteProperty "LoopFromLeft", LoopFromLeft, True
End Sub

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Sets the message to be displayed supports greek characters!"
Caption = Cptn
End Property
Public Property Let Caption(sCap As String)
Cptn = sCap
If Looped Then
Cptn = UCase(Cptn)
temp = String(Chars, " ") & Cptn
Else
Cptn = UCase(Cptn)
temp = String(Chars, " ") & Cptn & " " ' & String(leng, " ")
End If
PropertyChanged "Caption"
End Property

Public Property Get Interval() As Integer
Attribute Interval.VB_Description = "Sets the frequency of the scrolling"
Interval = Intrvl
End Property

Public Property Let Interval(sInt As Integer)
If sInt < 1 Then sInt = 1
If sInt > 2000 Then sInt = 2000
Intrvl = Int(sInt)
Timer1.Interval = Intrvl
PropertyChanged "Interval"
End Property

Public Property Get LoopFromLeft() As Boolean
Attribute LoopFromLeft.VB_Description = "Defines whether the message starts form the left or right"
LoopFromLeft = Looped
End Property
Public Property Let LoopFromLeft(sLoop As Boolean)
Looped = sLoop
If Looped Then
Cptn = UCase(Cptn)
temp = String(Chars, " ") & Cptn
Else
Cptn = UCase(Cptn)
temp = String(Chars, " ") & Cptn & " "
End If
PropertyChanged "LoopFromLeft"
End Property

Public Sub StartLoop()
Attribute StartLoop.VB_Description = "Starts the loop"
d = 0
Timer1.Enabled = True
End Sub

Public Sub StopLoop()
Attribute StopLoop.VB_Description = "Stops the loop"
Timer1.Enabled = False
For i = 1 To Chars
Select Case dig
Case 1
Image1(i).Picture = PicClip1.GraphicCell(46)
Case 2
Image1(i).Picture = PicClip2.GraphicCell(46)
Case 3
Image1(i).Picture = PicClip3.GraphicCell(46)
Case 4
Image1(i).Picture = PicClip4.GraphicCell(46)
End Select

Next
End Sub

Public Sub About()
Form1.Show 1
End Sub
