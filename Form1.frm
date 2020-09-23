VERSION 5.00
Object = "{50D47187-8271-11D4-8113-0000E87193A1}#14.0#0"; "Marquee.ocx"
Begin VB.Form Form1 
   BackColor       =   &H80000008&
   Caption         =   "Form1"
   ClientHeight    =   2070
   ClientLeft      =   450
   ClientTop       =   2280
   ClientWidth     =   11265
   LinkTopic       =   "Form1"
   ScaleHeight     =   2070
   ScaleWidth      =   11265
   Begin Scrolling_Marquee.Marquee Marquee1 
      Height          =   465
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   10920
      _ExtentX        =   19262
      _ExtentY        =   820
      DigitColor      =   3
      Caption         =   "ÈÅÁÍÙ ÅÉÓÁÉ ÐÏËÕ ÌÁ ÐÏËÕ ×ÏÍÔÑÇ"
      Interval        =   200
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Caption"
      Height          =   495
      Left            =   6210
      TabIndex        =   6
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "About"
      Height          =   495
      Left            =   7530
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Stop"
      Height          =   495
      Left            =   4290
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Left/Right"
      Height          =   495
      Left            =   4890
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Interval 100"
      Height          =   495
      Left            =   3570
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Colors"
      Height          =   495
      Left            =   2250
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   495
      Left            =   5610
      TabIndex        =   0
      Top             =   1440
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ooo
Private Sub Command1_Click()
Marquee1.StartLoop
End Sub

Private Sub Command2_Click()
ooo = ooo + 1
If ooo > 4 Then ooo = 1
Marquee1.Digitcolor = ooo
End Sub

Private Sub Command3_Click()
Marquee1.Interval = 100
End Sub


Private Sub Command4_Click()
If Marquee1.LoopFromLeft = True Then
Marquee1.LoopFromLeft = False
Else
Marquee1.LoopFromLeft = True
End If
End Sub

Private Sub Command5_Click()
Marquee1.StopLoop
End Sub

Private Sub Command6_Click()
Marquee1.About
End Sub

Private Sub Command7_Click()
Marquee1.Caption = "Ï ÃÅÍÉÊÏÓ ÄÅÉÊÔÇÓ ÔÏÕ ×ÑÇÌÁÔÉÓÔÇÑÉÏÕ ÓÇÌÅÑÁ ÁÍÅÂÇÊÅ ÊÁÔÁ 7,42%"
End Sub

Private Sub Form_Load()

End Sub

Private Sub Marquee1_GotFocus()

End Sub
