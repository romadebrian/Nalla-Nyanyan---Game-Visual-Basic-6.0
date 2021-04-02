VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nalla Nyanyanya"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12915
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   12915
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   14280
      Top             =   4440
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   14160
      Top             =   5400
   End
   Begin VB.Image Image3 
      Height          =   1200
      Left            =   15000
      Picture         =   "Form1.frx":0CCA
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   1200
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   340
      Width           =   855
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Score"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   340
      Width           =   735
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   240
      Picture         =   "Form1.frx":3882
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2160
   End
   Begin VB.Image CharNalla 
      Height          =   975
      Left            =   360
      Picture         =   "Form1.frx":3F42
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   1935
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   5760
      Visible         =   0   'False
      Width           =   5415
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   9551
      _cy             =   1085
   End
   Begin VB.Image Image1 
      Height          =   7125
      Left            =   0
      Picture         =   "Form1.frx":7219
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13665
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xpos As Single, ypos As Single
Dim score As Integer
Dim lokasi As Integer

Private Sub Form_Load()
Dim loc As String

loc = App.Path & "\test.gif"
'CharNalla.Navigate "about:" & "<html>" & "<body leftMargin=0 topMargin=0 marginheight=0 marginwidth=0 scroll=no>" _
    & "<img src=""" & loc & """></img></body></html>"
        
'xpos = (Me.ScaleWidth - CharNalla.Width) / 2
'ypos = (Me.ScaleHeight - CharNalla.Height) / 2
'CharNalla.Move xpos, ypos
    
CharNalla.Left = 0
CharNalla.Top = 1500
Image3.Top = 2500
Image3.Left = 15000
    
If Menu.Sound = "On" Then
    WindowsMediaPlayer1.settings.volume = 100
    WindowsMediaPlayer1.URL = (App.Path & "\Nyan Cat original.mp3")
    WindowsMediaPlayer1.settings.setMode "loop", True
Else
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyUp
        If CharNalla.Top <= 0 Then
        Else
            CharNalla.Top = CharNalla.Top - 500
        End If
                
    Case vbKeyRight
        If CharNalla.Left >= Me.ScaleWidth - CharNalla.Width Then
        Else
            CharNalla.Left = CharNalla.Left + 500
        End If
            
    Case vbKeyDown
        If CharNalla.Top >= Me.ScaleHeight - CharNalla.Height Then
        Else
            CharNalla.Top = CharNalla.Top + 500
        End If
            
    Case vbKeyLeft
        If CharNalla.Left <= 0 Then
            CharNalla.Left = 0
        Else
            CharNalla.Left = CharNalla.Left - 500
        End If
End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
Menu.Show
End Sub

Private Sub Timer1_Timer()
Image3.Left = Image3.Left - 500

Randomize
lokasi = Int((5 * Rnd) + 1)
    
End Sub

Private Sub Timer2_Timer()
If Image3.Left < 0 Then
    Image3.Left = 15000
End If
    
If (CharNalla.Left = Image3.Left) And CharNalla.Top = Image3.Top Then
    Image3.Left = 15000
    score = score + 1
    Label2.Caption = score
End If
End Sub
