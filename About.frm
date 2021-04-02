VERSION 5.00
Begin VB.Form About 
   BackColor       =   &H80000014&
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   6540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4290
   LinkTopic       =   "Form4"
   ScaleHeight     =   6540
   ScaleWidth      =   4290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
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
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   2760
      Left            =   120
      Picture         =   "About.frx":0000
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   3960
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Label8"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   2760
      Width           =   3495
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Label7"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   2400
      Width           =   3495
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Label6"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   5
      Top             =   2040
      Width           =   3495
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Gambar :"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Roma Debrian"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   3495
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Creator :"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Nalla Nyanyan"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Name :"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
