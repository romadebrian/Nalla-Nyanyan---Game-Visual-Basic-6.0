VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nalla Nyanyanya"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12915
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   12915
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Time_Animasi 
      Interval        =   100
      Left            =   11040
      Top             =   5520
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   11520
      Top             =   4560
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   12000
      Top             =   5520
   End
   Begin VB.Image KejuBusuk 
      Height          =   1200
      Index           =   2
      Left            =   9840
      Picture         =   "Form1.frx":0CCA
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   1200
   End
   Begin VB.Image KejuBusuk 
      Height          =   1200
      Index           =   1
      Left            =   6960
      Picture         =   "Form1.frx":35E8
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   1200
   End
   Begin VB.Image Keju 
      Height          =   1200
      Index           =   2
      Left            =   9720
      Picture         =   "Form1.frx":5F06
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   1200
   End
   Begin VB.Image Keju 
      Height          =   1200
      Index           =   1
      Left            =   11040
      Picture         =   "Form1.frx":8ABE
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1200
   End
   Begin VB.Image KejuBusuk 
      Height          =   1200
      Index           =   0
      Left            =   8520
      Picture         =   "Form1.frx":B676
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   1200
   End
   Begin VB.Image Keju 
      Height          =   1200
      Index           =   0
      Left            =   11400
      Picture         =   "Form1.frx":DF94
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   1200
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Picture         =   "Form1.frx":10B4C
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2160
   End
   Begin VB.Image CharNalla 
      Height          =   975
      Left            =   360
      Picture         =   "Form1.frx":1120C
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
      Picture         =   "Form1.frx":144E3
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
Dim score_poin As Integer
Dim loc As Integer
Dim hasil_loc As Integer
Dim jam
Dim tanggal
Dim pic As Integer

Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Private Sub Form_Load()
Dim loc As String

pic = 1

Menu.status_form_main = 1
score_poin = 0

'loc = App.Path & "\test.gif"
'CharNalla.Navigate "about:" & "<html>" & "<body leftMargin=0 topMargin=0 marginheight=0 marginwidth=0 scroll=no>" _
    & "<img src=""" & loc & """></img></body></html>"
        
'xpos = (Me.ScaleWidth - CharNalla.Width) / 2
'ypos = (Me.ScaleHeight - CharNalla.Height) / 2
'CharNalla.Move xpos, ypos
    
CharNalla.Left = 0
CharNalla.Top = 1500

For i = 0 To Keju.Count - 1
    Lokasi
    Keju(i).Top = hasil_loc
    Keju(0).Left = 13000
    Keju(1).Left = 14000
    Keju(2).Left = 15000
Next

For i = 0 To KejuBusuk.Count - 1
    Lokasi
    KejuBusuk(i).Top = hasil_loc
    KejuBusuk(0).Left = 15000
    KejuBusuk(1).Left = 16000
    KejuBusuk(2).Left = 17000
Next
    
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
status_form = 0
Menu.Show
End Sub

Private Sub Time_Animasi_Timer()
pic = pic + 1

'pakai array dan ubah nama file jadi 123456

Select Case pic
    Case 1
        CharNalla.Picture = LoadPicture(App.Path & "\CatNalla\1.gif")
    Case 2
        CharNalla.Picture = LoadPicture(App.Path & "\CatNalla\2.gif")
    Case 3
        CharNalla.Picture = LoadPicture(App.Path & "\CatNalla\3.gif")
    Case 4
        CharNalla.Picture = LoadPicture(App.Path & "\CatNalla\4.gif")
    Case 5
        CharNalla.Picture = LoadPicture(App.Path & "\CatNalla\5.gif")
    Case 6
        CharNalla.Picture = LoadPicture(App.Path & "\CatNalla\6.gif")
        pic = 1
End Select
End Sub

Private Sub Timer1_Timer()
For i = 0 To Keju.Count - 1
    Keju(i).Left = Keju(i).Left - 500
Next

For i = 0 To Keju.Count - 1
    If Keju(i).Left < 0 Then
        Keju(i).Left = 13000
        Lokasi
        Keju(i).Top = hasil_loc
    End If
Next

'=====================================================

For i = 0 To KejuBusuk.Count - 1
    KejuBusuk(i).Left = KejuBusuk(i).Left - 500
Next

For i = 0 To KejuBusuk.Count - 1
    If KejuBusuk(i).Left < 0 Then
        KejuBusuk(i).Left = 13000
        Lokasi
        KejuBusuk(i).Top = hasil_loc
    End If
Next
End Sub

Private Sub Timer2_Timer()
For i = 0 To Keju.Count - 1
    If (CharNalla.Left = Keju(i).Left) And CharNalla.Top = Keju(i).Top Then
        Keju(i).Left = 13000
        score_poin = score_poin + 10
        Label2.Caption = score_poin
        
        sndPlaySound App.Path & "\hit.wav", 1 Or 2
        Lokasi
        Keju(i).Top = hasil_loc
    End If
Next

For i = 0 To KejuBusuk.Count - 1
    If (CharNalla.Left = KejuBusuk(i).Left) And CharNalla.Top = KejuBusuk(i).Top Then
    'MsgBox "Game Over"
    Timer1.Enabled = False
    Timer2.Enabled = False
    
    Score.List1.Clear
    Score.List1.AddItem Label2.Caption
    Score.simpan_score
    
    jam = Time
    tanggal = Date
    
    Score.List2.Clear
    Score.List2.AddItem Time + Date
    Score.simpan_waktu
    
    GameOver.Label2.Caption = score_poin
    GameOver.Show
    End If
Next
End Sub

Sub Lokasi()
Randomize
loc = Int((10 * Rnd) + 1)

Select Case loc
Case 1
    hasil_loc = 2000
Case 2
    hasil_loc = 2500
Case 3
    hasil_loc = 3000
Case 4
    hasil_loc = 3500
Case 5
    hasil_loc = 4000
Case 6
    hasil_loc = 4500
Case 7
    hasil_loc = 5000
Case 8
    hasil_loc = 5500
Case 9
    hasil_loc = 6000
Case 10
    hasil_loc = 6500
Case Else
    hasil_loc = 7000
End Select
End Sub

Sub Tambah_Keju_Biasa()
Dim k As Integer

k = KejuBusuk.Count
Load KejuBusuk(k)
KejuBusuk(k).Visible = True

End Sub




























































