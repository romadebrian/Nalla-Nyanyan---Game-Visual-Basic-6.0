VERSION 5.00
Begin VB.Form About 
   BackColor       =   &H80000014&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4170
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   4170
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Nyanyanyanyanyanyanya! by daniwellP"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   1920
      Width           =   3495
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Music (BGM)"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1440
      TabIndex        =   8
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   2760
      Left            =   120
      Picture         =   "About.frx":0CCA
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   3960
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "@Rizukiaeru (Gambar Nalla di Sebelaj Kiri Menu)"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   3360
      Width           =   3495
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Nalla (Gambar Nalla di Sebelah Kanan Menu)"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   3000
      Width           =   3495
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "athayamansion (Main Character)"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   5
      Top             =   2640
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
      Top             =   2280
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
      Top             =   1200
      Width           =   3495
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Creator"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Nalla Nyanyanya (Versi Beta)"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Name"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()
MsgBox "Saya minta maaf jika salah mencantumkan nama illustrator"
End Sub
