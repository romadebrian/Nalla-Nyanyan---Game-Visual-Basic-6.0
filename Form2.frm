VERSION 5.00
Begin VB.Form Menu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nalla Nyanyanya"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13470
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   13700.4
   ScaleMode       =   0  'User
   ScaleWidth      =   23033.45
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3480
      Top             =   600
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "BGM OFF"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   615
      Left            =   6120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Image Image9 
      Height          =   840
      Left            =   120
      Picture         =   "Form2.frx":0CCA
      Stretch         =   -1  'True
      Top             =   195
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image Image8 
      Height          =   780
      Left            =   120
      Picture         =   "Form2.frx":24DE
      Stretch         =   -1  'True
      Top             =   240
      Width           =   975
   End
   Begin VB.Image Image7 
      Height          =   4935
      Left            =   9480
      Picture         =   "Form2.frx":2F12
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   4215
   End
   Begin VB.Image Image6 
      Height          =   5400
      Left            =   -240
      Picture         =   "Form2.frx":477B4
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   5160
   End
   Begin VB.Image Image5 
      Height          =   885
      Left            =   4680
      Picture         =   "Form2.frx":75D05
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   4305
   End
   Begin VB.Image Image4 
      Height          =   885
      Left            =   4680
      Picture         =   "Form2.frx":7710E
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   4305
   End
   Begin VB.Image Image3 
      Height          =   885
      Left            =   4680
      Picture         =   "Form2.frx":78608
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   4305
   End
   Begin VB.Image Image2 
      Height          =   1215
      Left            =   4200
      Picture         =   "Form2.frx":79A8F
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   5295
   End
   Begin VB.Image Image1 
      Height          =   7155
      Left            =   0
      Picture         =   "Form2.frx":7B6CB
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13498
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Option Explicit
Public Sound
Public status_form_main As Integer

Private Sub Form_Load()
Sound = "On"
status_form_main = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Image2_Click()
Main.Show
Menu.Hide
End Sub

Private Sub Image3_Click()
Score.Show
End Sub

Private Sub Image4_Click()
About.Show
End Sub

Private Sub Image5_Click()
End
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Picture = LoadPicture(App.Path & "\Interface\Play2.gif")
End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Picture = LoadPicture(App.Path & "\Interface\Play.gif")
End Sub

Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Picture = LoadPicture(App.Path & "\Interface\Score2.gif")
End Sub

Private Sub Image3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Picture = LoadPicture(App.Path & "\Interface\Score.gif")
End Sub

Private Sub Image4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image4.Picture = LoadPicture(App.Path & "\Interface\About2.gif")
End Sub

Private Sub Image4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image4.Picture = LoadPicture(App.Path & "\Interface\About.gif")
End Sub

Private Sub Image5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image5.Picture = LoadPicture(App.Path & "\Interface\Exit2.gif")
End Sub

Private Sub Image5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image5.Picture = LoadPicture(App.Path & "\Interface\Exit.gif")
End Sub

Private Sub Image8_Click()
Sound = "Off"
Image9.Visible = True
Image8.Visible = False

Label1.Caption = "BGM OFF"
Label1.Visible = True
Timer1.Enabled = True
End Sub

Private Sub Image9_Click()
Sound = "On"
Image8.Visible = True
Image9.Visible = False

Label1.Caption = "BGM ON"
Label1.Visible = True
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
Label1.Visible = False
End Sub
