VERSION 5.00
Begin VB.Form Score 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   7410
   ClientLeft      =   8235
   ClientTop       =   4050
   ClientWidth     =   6825
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   6825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6840
      Width           =   975
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      Height          =   3735
      Left            =   3360
      TabIndex        =   3
      Top             =   3000
      Width           =   2000
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   3735
      Left            =   1440
      TabIndex        =   0
      Top             =   3000
      Width           =   2000
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C000C0&
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   2640
      Width           =   1995
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C000C0&
      Caption         =   "Score"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   2640
      Width           =   1995
   End
   Begin VB.Image Image1 
      Height          =   7575
      Left            =   -120
      Picture         =   "Score.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7100
   End
End
Attribute VB_Name = "Score"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Menu.status_form_main = 1 Then
    Unload Main
    Unload Score
Else
    Unload Me
End If
End Sub

Private Sub Form_Load()
Dim ff As Long
Dim line As String

List1.Clear

ff = FreeFile
Open App.Path & "\score.txt" For Input As #ff
Do While Not EOF(ff)
       Line Input #ff, line
       'make sure we're not adding a blank line
       If Len(line) Then List1.AddItem line
Loop
Close #ff
End Sub

Sub simpan_score()
Dim i As Integer

Open App.Path & "\score.txt" For Append As #1 'Drive penyimpanan

For i = 0 To List1.ListCount - 1
    Print #1, List1.List(i)
Next
Close #1
End Sub
