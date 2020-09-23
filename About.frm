VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Digital Marquee"
   ClientHeight    =   1395
   ClientLeft      =   2385
   ClientTop       =   2505
   ClientWidth     =   5025
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   5025
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   120
      Top             =   1080
   End
   Begin Scrolling_Marquee.Marquee Marquee1 
      Height          =   465
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4620
      _ExtentX        =   8149
      _ExtentY        =   820
      DigitColor      =   3
      Caption         =   $"About.frx":030A
      Interval        =   150
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Digital Marquee By Charitakis Panagiotis E-mail:panos000@yahoo.com"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   4575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim k As Integer
Private Sub Form_Load()
Marquee1.StartLoop
End Sub

Private Sub Timer1_Timer()
k = k + 1
If k > 4 Then k = 1
Marquee1.Digitcolor = k
Select Case k
Case 1
Label1.ForeColor = vbCyan
Case 2
Label1.ForeColor = vbGreen
Case 3
Label1.ForeColor = vbRed
Case 4
Label1.ForeColor = vbYellow
End Select
End Sub
