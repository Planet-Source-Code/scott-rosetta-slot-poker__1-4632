VERSION 5.00
Begin VB.Form SlotPoker 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Slot Poker"
   ClientHeight    =   2775
   ClientLeft      =   5865
   ClientTop       =   4185
   ClientWidth     =   3735
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Xit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   2280
      TabIndex        =   5
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton helllllp 
      Caption         =   "Help"
      Height          =   495
      Left            =   2280
      TabIndex        =   4
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton Pick 
      Caption         =   "Spin!"
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   2175
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   2520
      Picture         =   "SlotPoker.frx":0000
      ScaleHeight     =   1455
      ScaleWidth      =   1095
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   1320
      Picture         =   "SlotPoker.frx":06F3
      ScaleHeight     =   1455
      ScaleWidth      =   1095
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   120
      Picture         =   "SlotPoker.frx":0DE6
      ScaleHeight     =   1455
      ScaleWidth      =   1095
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "SlotPoker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1
Dim Random As Integer
Dim Random2 As Integer
Dim Random3 As Integer
Dim I As Integer

Private Sub helllllp_Click()
MsgBox ("Really a no brainer here... Click spin and see what pops up!")
End Sub

Private Sub Pick_Click()

FirstRoll:
    Picture1.Picture = LoadPicture("ace-diamond.gif")
    Picture2.Picture = LoadPicture("ace-diamond.gif")
    Picture3.Picture = LoadPicture("ace-diamond.gif")
    Picture1.Picture = LoadPicture("two-diamond.gif")
    Picture2.Picture = LoadPicture("two-diamond.gif")
    Picture3.Picture = LoadPicture("two-diamond.gif")
    Picture1.Picture = LoadPicture("three-diamond.gif")
    Picture1.Picture = LoadPicture("cardback.gif")
    Picture2.Picture = LoadPicture("three-diamond.gif")
    Picture3.Picture = LoadPicture("three-diamond.gif")
    Picture1.Picture = LoadPicture("four-diamond.gif")
    Picture2.Picture = LoadPicture("four-diamond.gif")
    Picture3.Picture = LoadPicture("four-diamond.gif")
    Picture1.Picture = LoadPicture("five-diamond.gif")
    Picture2.Picture = LoadPicture("five-diamond.gif")
    Picture3.Picture = LoadPicture("five-diamond.gif")
    Picture1.Picture = LoadPicture("six-diamond.gif")
    Picture2.Picture = LoadPicture("six-diamond.gif")
    Picture3.Picture = LoadPicture("six-diamond.gif")
    Picture1.Picture = LoadPicture("seven-diamond.gif")
    Picture3.Picture = LoadPicture("cardback.gif")
    Picture2.Picture = LoadPicture("seven-diamond.gif")
    Picture3.Picture = LoadPicture("seven-diamond.gif")
    Picture1.Picture = LoadPicture("eight-diamond.gif")
    Picture2.Picture = LoadPicture("eight-diamond.gif")
    Picture3.Picture = LoadPicture("eight-diamond.gif")
    Picture1.Picture = LoadPicture("nine-diamond.gif")
    Picture2.Picture = LoadPicture("nine-diamond.gif")
    Picture3.Picture = LoadPicture("nine-diamond.gif")
    Picture1.Picture = LoadPicture("ten-diamond.gif")
    Picture2.Picture = LoadPicture("ten-diamond.gif")
    Picture3.Picture = LoadPicture("ten-diamond.gif")
    Picture1.Picture = LoadPicture("jack-diamond.gif")
    Picture2.Picture = LoadPicture("jack-diamond.gif")
    Picture3.Picture = LoadPicture("jack-diamond.gif")
    Picture1.Picture = LoadPicture("queen-diamond.gif")
    Picture2.Picture = LoadPicture("cardback.gif")
    Picture2.Picture = LoadPicture("queen-diamond.gif")
    Picture3.Picture = LoadPicture("queen-diamond.gif")
    Picture1.Picture = LoadPicture("king-diamond.gif")
    Picture2.Picture = LoadPicture("king-diamond.gif")
    Picture3.Picture = LoadPicture("king-diamond.gif")
    GoTo FirstPick:

FirstPick:
    For I = 1 To 1
        Randomize
        Random = Int((Rnd * 13) + 1)
        If Random = 1 Then Picture1.Picture = LoadPicture("ace-diamond.gif")
        If Random = 2 Then Picture1.Picture = LoadPicture("two-diamond.gif")
        If Random = 3 Then Picture1.Picture = LoadPicture("three-diamond.gif")
        If Random = 4 Then Picture1.Picture = LoadPicture("four-diamond.gif")
        If Random = 5 Then Picture1.Picture = LoadPicture("five-diamond.gif")
        If Random = 6 Then Picture1.Picture = LoadPicture("six-diamond.gif")
        If Random = 7 Then Picture1.Picture = LoadPicture("seven-diamond.gif")
        If Random = 8 Then Picture1.Picture = LoadPicture("eight-diamond.gif")
        If Random = 9 Then Picture1.Picture = LoadPicture("nine-diamond.gif")
        If Random = 10 Then Picture1.Picture = LoadPicture("ten-diamond.gif")
        If Random = 11 Then Picture1.Picture = LoadPicture("jack-diamond.gif")
        If Random = 12 Then Picture1.Picture = LoadPicture("queen-diamond.gif")
        If Random = 13 Then Picture1.Picture = LoadPicture("king-diamond.gif")
    Next I
    GoTo SecondRoll:
      
SecondRoll:
    Picture2.Picture = LoadPicture("ace-diamond.gif")
    Picture3.Picture = LoadPicture("ace-diamond.gif")
    Picture2.Picture = LoadPicture("two-diamond.gif")
    Picture3.Picture = LoadPicture("two-diamond.gif")
    Picture2.Picture = LoadPicture("three-diamond.gif")
    Picture3.Picture = LoadPicture("three-diamond.gif")
    Picture2.Picture = LoadPicture("four-diamond.gif")
    Picture3.Picture = LoadPicture("four-diamond.gif")
    Picture2.Picture = LoadPicture("five-diamond.gif")
    Picture3.Picture = LoadPicture("five-diamond.gif")
    Picture2.Picture = LoadPicture("six-diamond.gif")
    Picture3.Picture = LoadPicture("six-diamond.gif")
    Picture3.Picture = LoadPicture("cardback.gif")
    Picture2.Picture = LoadPicture("seven-diamond.gif")
    Picture3.Picture = LoadPicture("seven-diamond.gif")
    Picture2.Picture = LoadPicture("eight-diamond.gif")
    Picture3.Picture = LoadPicture("eight-diamond.gif")
    Picture2.Picture = LoadPicture("nine-diamond.gif")
    Picture3.Picture = LoadPicture("nine-diamond.gif")
    Picture2.Picture = LoadPicture("ten-diamond.gif")
    Picture3.Picture = LoadPicture("ten-diamond.gif")
    Picture2.Picture = LoadPicture("jack-diamond.gif")
    Picture3.Picture = LoadPicture("jack-diamond.gif")
    Picture2.Picture = LoadPicture("cardback.gif")
    Picture2.Picture = LoadPicture("queen-diamond.gif")
    Picture3.Picture = LoadPicture("queen-diamond.gif")
    Picture2.Picture = LoadPicture("king-diamond.gif")
    Picture3.Picture = LoadPicture("king-diamond.gif")
    GoTo SecondPick:

SecondPick:
    For I = 1 To 1
        Randomize
        Random2 = Int((Rnd * 13) + 1)
        If Random2 = 1 Then Picture2.Picture = LoadPicture("ace-diamond.gif")
        If Random2 = 2 Then Picture2.Picture = LoadPicture("two-diamond.gif")
        If Random2 = 3 Then Picture2.Picture = LoadPicture("three-diamond.gif")
        If Random2 = 4 Then Picture2.Picture = LoadPicture("four-diamond.gif")
        If Random2 = 5 Then Picture2.Picture = LoadPicture("five-diamond.gif")
        If Random2 = 6 Then Picture2.Picture = LoadPicture("six-diamond.gif")
        If Random2 = 7 Then Picture2.Picture = LoadPicture("seven-diamond.gif")
        If Random2 = 8 Then Picture2.Picture = LoadPicture("eight-diamond.gif")
        If Random2 = 9 Then Picture2.Picture = LoadPicture("nine-diamond.gif")
        If Random2 = 10 Then Picture2.Picture = LoadPicture("ten-diamond.gif")
        If Random2 = 11 Then Picture2.Picture = LoadPicture("jack-diamond.gif")
        If Random2 = 12 Then Picture2.Picture = LoadPicture("queen-diamond.gif")
        If Random2 = 13 Then Picture2.Picture = LoadPicture("king-diamond.gif")
    Next I
    GoTo ThirdRoll:
    
ThirdRoll:
    Picture3.Picture = LoadPicture("ace-diamond.gif")
    Picture3.Picture = LoadPicture("two-diamond.gif")
    Picture3.Picture = LoadPicture("three-diamond.gif")
    Picture3.Picture = LoadPicture("four-diamond.gif")
    Picture3.Picture = LoadPicture("five-diamond.gif")
    Picture3.Picture = LoadPicture("six-diamond.gif")
    Picture3.Picture = LoadPicture("cardback.gif")
    Picture3.Picture = LoadPicture("seven-diamond.gif")
    Picture3.Picture = LoadPicture("eight-diamond.gif")
    Picture3.Picture = LoadPicture("nine-diamond.gif")
    Picture3.Picture = LoadPicture("ten-diamond.gif")
    Picture3.Picture = LoadPicture("jack-diamond.gif")
    Picture3.Picture = LoadPicture("queen-diamond.gif")
    Picture3.Picture = LoadPicture("king-diamond.gif")
    GoTo ThirdPick:
    
ThirdPick:
    For I = 1 To 1
        Randomize
        Random3 = Int((Rnd * 13) + 1)
        If Random3 = 1 Then Picture3.Picture = LoadPicture("ace-diamond.gif")
        If Random3 = 2 Then Picture3.Picture = LoadPicture("two-diamond.gif")
        If Random3 = 3 Then Picture3.Picture = LoadPicture("three-diamond.gif")
        If Random3 = 4 Then Picture3.Picture = LoadPicture("four-diamond.gif")
        If Random3 = 5 Then Picture3.Picture = LoadPicture("five-diamond.gif")
        If Random3 = 6 Then Picture3.Picture = LoadPicture("six-diamond.gif")
        If Random3 = 7 Then Picture3.Picture = LoadPicture("seven-diamond.gif")
        If Random3 = 8 Then Picture3.Picture = LoadPicture("eight-diamond.gif")
        If Random3 = 9 Then Picture3.Picture = LoadPicture("nine-diamond.gif")
        If Random3 = 10 Then Picture3.Picture = LoadPicture("ten-diamond.gif")
        If Random3 = 11 Then Picture3.Picture = LoadPicture("jack-diamond.gif")
        If Random3 = 12 Then Picture3.Picture = LoadPicture("queen-diamond.gif")
        If Random3 = 13 Then Picture3.Picture = LoadPicture("king-diamond.gif")
    Next I

End Sub

Private Sub Xit_Click()
Unload Me
End Sub
