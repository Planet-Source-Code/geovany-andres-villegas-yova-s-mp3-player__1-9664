VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Yova's MPlayer"
   ClientHeight    =   6210
   ClientLeft      =   1170
   ClientTop       =   1530
   ClientWidth     =   10305
   Icon            =   "Player.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   414
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   687
   Begin VB.CommandButton Command2 
      Caption         =   ">\"
      Height          =   375
      Left            =   4560
      TabIndex        =   11
      Top             =   4200
      Width           =   1095
   End
   Begin VB.TextBox TiempO 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   360
      Left            =   4725
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "00:00"
      Top             =   3360
      Width           =   645
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1800
      Top             =   4800
   End
   Begin MSComDlg.CommonDialog CD2 
      Left            =   4920
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "YovasMP3Player's Lista [*.ply]|*.ply|"
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "OPCIONES DE LISTA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1380
      Left            =   7320
      TabIndex        =   9
      Top             =   1695
      Width           =   2415
      Begin VB.Image Image14 
         Height          =   345
         Left            =   435
         Top             =   840
         Width           =   1530
      End
      Begin VB.Image Image13 
         Height          =   345
         Left            =   360
         Top             =   360
         Width           =   1680
      End
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   1440
      Top             =   4800
   End
   Begin VB.Timer Timer4 
      Interval        =   1
      Left            =   1080
      Top             =   4800
   End
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   720
      Top             =   4800
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   360
      Top             =   4800
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   4800
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   375
      Left            =   3000
      TabIndex        =   7
      Top             =   3840
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000012&
      Caption         =   "BALANCE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   855
      Left            =   7560
      TabIndex        =   5
      Top             =   3240
      Width           =   2175
      Begin MSComctlLib.Slider Slider3 
         Height          =   315
         Left            =   270
         TabIndex        =   6
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   1
         Min             =   -5000
         Max             =   5000
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000007&
         Caption         =   "CENTRO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   720
         TabIndex        =   8
         Top             =   600
         Width           =   795
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000007&
      Caption         =   "VOLUMEN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   855
      Left            =   240
      TabIndex        =   3
      Top             =   3240
      Width           =   2175
      Begin MSComctlLib.Slider Slider2 
         Height          =   320
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   1
         Max             =   2500
      End
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1530
      Left            =   2760
      TabIndex        =   1
      Top             =   1800
      Width           =   4455
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   5400
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Archivos de Sonido[*.mp3] [*.wav]|*.mp3;*.wav|"
   End
   Begin VB.Image Image12 
      Height          =   345
      Left            =   2880
      ToolTipText     =   "Activar reproducción al azar"
      Top             =   1320
      Width           =   1440
   End
   Begin VB.Image Image9 
      Height          =   345
      Left            =   5850
      ToolTipText     =   "Activar repetición"
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Image Image7 
      Height          =   405
      Left            =   1080
      ToolTipText     =   "Borrar lista de reproducción"
      Top             =   2520
      Width           =   1710
   End
   Begin VB.Image Image8 
      Height          =   405
      Left            =   1080
      ToolTipText     =   "Remover archivo"
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Image Image6 
      Height          =   345
      Left            =   4365
      ToolTipText     =   "Desactivar reproducción continua"
      Top             =   1320
      Width           =   1440
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   570
      Left            =   2280
      TabIndex        =   2
      Top             =   4530
      Width           =   5355
   End
   Begin VB.Image Image5 
      Height          =   405
      Left            =   4200
      Top             =   120
      Width           =   1695
   End
   Begin VB.Image Image4 
      Height          =   405
      Left            =   7320
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Image Image3 
      Height          =   405
      Left            =   5040
      Top             =   720
      Width           =   1980
   End
   Begin VB.Image Image2 
      Height          =   405
      Left            =   3120
      Top             =   720
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   405
      Left            =   1140
      Top             =   1200
      Width           =   1695
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   735
      Left            =   6000
      TabIndex        =   0
      Top             =   5400
      Width           =   615
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   0
      WindowlessVideo =   0   'False
   End
   Begin VB.Menu madd 
      Caption         =   "Add"
      Begin VB.Menu mfile 
         Caption         =   "Archivo"
         Shortcut        =   ^A
      End
      Begin VB.Menu mdirectory 
         Caption         =   "Directorio"
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu mpop 
      Caption         =   "Pop"
      Begin VB.Menu mminimize 
         Caption         =   "Minimizar"
         Shortcut        =   ^W
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu msearch 
         Caption         =   "Buscar"
         Shortcut        =   ^B
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mabout 
         Caption         =   "Acercade Yova's MPlayer"
         Shortcut        =   ^Y
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu montop 
         Caption         =   "Siempre visible"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MoverForm As Boolean
Public Playing As Boolean
Public Entrada As Boolean
Public Salida As Boolean
Public SigPlay As Boolean
Public SeDetuvo As Boolean
Public Continuar As Boolean
Public Replay As Boolean
Public Shuffle As Boolean
Public Pausa As Boolean
Public Paro As Boolean
Public Volumen1 As Boolean
Public FirstImage As Single
Public SecondImage As Single
Public ThirdImage As Single
Public Label As String
Public ListName As String
Public Eslabon As Long
Dim Sig(10) As Integer

Private Sub Command2_Click()
irSiguiente
End Sub

Private Sub Form_Load()
 Dim Region1 As Long
    Dim Region2 As Long
    Dim CombinarRgn As Long

    
    Region1 = CreateRectRgn(201, 160, 473, 275)
    Region2 = CreateEllipticRgn(201, 280, 473, 350)
    CombinarRgn = CombineRgn(Region1, Region1, Region2, RGN_OR)
    Region2 = CreateEllipticRgn(5, 245, 182, 335)
    CombinarRgn = CombineRgn(Region1, Region1, Region2, RGN_OR)
    Region2 = CreateRoundRectRgn(150, 330, 520, 380, 80, 200) 'Region titulo de la cancion
    CombinarRgn = CombineRgn(Region1, Region1, Region2, RGN_OR)
    Region2 = CreateEllipticRgn(490, 248, 675, 333)
    CombinarRgn = CombineRgn(Region1, Region1, Region2, RGN_OR)
    Region2 = CreateEllipticRgn(20, 70, 675, 333)
    CombinarRgn = CombineRgn(Region1, Region1, Region2, RGN_OR)
    Region2 = CreateEllipticRgn(200, 45, 475, 203)
    CombinarRgn = CombineRgn(Region1, Region1, Region2, RGN_OR)
    Region2 = CreateRectRgn(320, 268, 359, 288) 'Region reloj
    CombinarRgn = CombineRgn(Region1, Region1, Region2, RGN_DIFF)
    Region2 = CreateRectRgn(322, 270, 357, 286) 'Region reloj
    CombinarRgn = CombineRgn(Region1, Region1, Region2, RGN_OR)
    CombinarRgn = SetWindowRgn(Me.hWnd, Region1, True)

Call Slider2_Scroll
Pausa = False
Paro = True
Playing = False
Volumen1 = True
Continuar = True
CD2.InitDir = (App.Path + "\")
Entrada = True
LoadList

    If List1.ListCount > 0 Then
    List1.Selected(0) = True
    TxT = List1.List(List1.ListIndex)
    TenerNombre
    End If

CargarImagenes

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MDown Button, X, Y
        
    If Button = 2 Then
        Form1.PopupMenu mpop
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MMove Button, X, Y
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 1 Then
If MoverForm Then
MoverForm = False
Form1.MousePointer = 0
End If
End If
End Sub





Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Image1.Picture = LoadPicture(App.Path + "\Pix\abrir_on.jpg")
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Image1.Picture = LoadPicture(App.Path + "\Pix\abrir_off.jpg")
Form1.PopupMenu madd, , 75, 88
End If
End Sub

Private Sub Image12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    If Image12.Picture = ThirdImage Then
        Timer4.Enabled = False
        Image12.Picture = LoadPicture(App.Path + "\Pix\shuffle_on.jpg")
        Image12.ToolTipText = "Desactivar reproducción al azar"
        Shuffle = True
        Namecito = ""
    Else
        Image12.Picture = LoadPicture(App.Path + "\Pix\shuffle_off.jpg")
        Image12.ToolTipText = "Activar reproducción al azar"
        Timer4.Enabled = True
        Shuffle = False
    End If
End If
End Sub


Private Sub Image13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Image13.Picture = LoadPicture(App.Path + "\Pix\save_on.jpg")
End Sub

Private Sub Image13_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Image13.Picture = LoadPicture(App.Path + "\Pix\save_off.jpg")
If List1.ListCount > 0 Then
    Seguro
Else
    MsgBox "No hay elementos en la lista de reproducción", vbOKOnly + vbExclamation, "Error"
End If

End If
End Sub
Private Sub Image14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Image14.Picture = LoadPicture(App.Path + "\Pix\load_on.jpg")
End Sub

Private Sub Image14_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Image14.Picture = LoadPicture(App.Path + "\Pix\load_off.jpg")
CD2.FileName = ""
CD2.ShowOpen
ListName = CD2.FileName
    If ListName <> "" Then
        LoadList
                If List1.ListCount > 0 Then
                        For j = 0 To (List1.ListCount - 1)
                            If List1.List(j) <> "" Then
                            List1.Selected(j) = True
                            TxT = List1.Text
                            TenerNombre
                            Exit Sub
                            End If
                        Next j
                End If
    End If

End If
End Sub
Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Image2.Picture = LoadPicture(App.Path + "\Pix\play_on.jpg")
End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
MUPlay Button, X, Y
End If
End Sub


Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Image3.Picture = LoadPicture(App.Path + "\Pix\pausa_on.jpg")
End Sub

Private Sub Image3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Playing = False Then
        If Paro = False Then
        Image3.Picture = LoadPicture(App.Path + "\Pix\pausa_off.jpg")
        Timer1.Enabled = True
        MediaPlayer1.Play
        Timer6.Enabled = True
        Image2.Picture = LoadPicture(App.Path + "\Pix\play_on.jpg")
        Playing = True
        Else
        Image3.Picture = LoadPicture(App.Path + "\Pix\pausa_off.jpg")
        Exit Sub
        End If
Else
Image2.Picture = LoadPicture(App.Path + "\Pix\play_off.jpg")
Image3.Picture = LoadPicture(App.Path + "\Pix\pausa_on.jpg")
MediaPlayer1.Pause
Timer1.Enabled = False
Timer6.Enabled = False
Playing = False
Pausa = True
End If

End Sub

Private Sub Image4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Image4.Picture = LoadPicture(App.Path + "\Pix\stop_on.jpg")
End Sub

Private Sub Image4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Image4.Picture = LoadPicture(App.Path + "\Pix\stop_off.jpg")
Image2.Picture = LoadPicture(App.Path + "\Pix\play_off.jpg")
Image3.Picture = LoadPicture(App.Path + "\Pix\pausa_off.jpg")
MediaPlayer1.Stop
MediaPlayer1.FileName = ""
Playing = False
Paro = True
Timer6.Enabled = False
TiempO.Text = "00:00"
Slider1.Value = 0
Timer5.Enabled = False
TxT = List1.List(List1.ListIndex)
TenerNombre
End If
End Sub

Private Sub Image5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Image5.Picture = LoadPicture(App.Path + "\Pix\salir_on.jpg")
End If
End Sub



Private Sub Image5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim retorno As String

If Button = 1 Then
Image5.Picture = LoadPicture(App.Path + "\Pix\salir_off.jpg")
retorno = MsgBox("En realidad quieres salir?", vbOKCancel + vbQuestion, "Salir?")
    If retorno = vbOK Then
    Salida = True
    SaveList
    End
    End If
End If
End Sub

Private Sub MDown(Button As Integer, xx As Single, yy As Single)

If Button = 1 Then
        If Not MoverForm Then
        Difx = xx
        Dify = yy
        Form1.MousePointer = 15
        MoverForm = True
        End If
End If
End Sub

Private Sub MMove(Buton As Integer, X As Single, Y As Single)

    If Buton = 1 Then
        If MoverForm = True Then
            Move Left + (X - Difx), Top + (Y - Dify)
            Form1.Refresh
            DoEvents
            End If
    End If

End Sub

Private Sub Image6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    If Image6.Picture = FirstImage Then
        Timer2.Enabled = False
        Image6.Picture = LoadPicture(App.Path + "\Pix\cont_on.jpg")
        Image6.ToolTipText = "Desactivar reproducción continua"
        Continuar = True
    Else
        Image6.Picture = LoadPicture(App.Path + "\Pix\cont_off.jpg")
        Image6.ToolTipText = "Activar reproducción continua"
        Continuar = False
        Timer2.Enabled = True
    End If
End If
End Sub



Private Sub Image7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Image7.Picture = LoadPicture(App.Path + "\Pix\clear_on.jpg")
End Sub

Private Sub Image7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Image7.Picture = LoadPicture(App.Path + "\Pix\clear_off.jpg")
        
        If List1.ListCount > 0 Then
            List1.Clear
            Slider1.SetFocus
            Image4.Picture = LoadPicture(App.Path + "\Pix\stop_off.jpg")
            Image2.Picture = LoadPicture(App.Path + "\Pix\play_off.jpg")
            Image3.Picture = LoadPicture(App.Path + "\Pix\pausa_off.jpg")
            MediaPlayer1.Stop
            MediaPlayer1.FileName = ""
            Label1.Caption = ""
            Namecito = ""
            Playing = False
            Timer1.Enabled = False
            Timer6.Enabled = False
            Paro = True
            TiempO.Text = "00:00"
            Slider1.Value = 0
        Else
            MsgBox "No hay ningún elemento en la lista de reproducción", vbOKOnly + vbExclamation, "Error"
        End If
        
End If
End Sub
Private Sub Image8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Image8.Picture = LoadPicture(App.Path + "\Pix\rem_on.jpg")
End Sub

Private Sub Image8_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Lugar As Integer
Dim Volvi As Integer

If Button = 1 Then

Lugar = (List1.ListIndex - 1)
Image8.Picture = LoadPicture(App.Path + "\Pix\rem_off.jpg")

If List1.ListCount > 0 Then
            If List1.ListIndex >= 0 Then
                        If List1.List(List1.ListIndex) = MediaPlayer1.FileName Then
                        Image4.Picture = LoadPicture(App.Path + "\Pix\stop_off.jpg")
                        Image2.Picture = LoadPicture(App.Path + "\Pix\play_off.jpg")
                        Image3.Picture = LoadPicture(App.Path + "\Pix\pausa_off.jpg")
                        MediaPlayer1.Stop
                        MediaPlayer1.FileName = ""
                        Playing = False
                        Paro = True
                        Timer1.Enabled = False
                        Timer6.Enabled = False
                        TiempO.Text = "00:00"
                        Slider1.Value = 0
                        End If
                        '-
                                            If Namecito <> "" Then
                                                    Nameci = Left(Namecito, (Len(Namecito) - 1))
                                                        Usados = Split(Nameci, "|")
                                                            For i = LBound(Usados) To UBound(Usados)
                                                                        For m = 0 To (List1.ListCount - 1)
                                                                            If Usados(i) = List1.List(List1.ListIndex) Then
                                                                                Usados(i) = ""
                                                                            End If
                                                                        Next m
                                                            Next i
                                            End If
                        '-
                 List1.RemoveItem List1.ListIndex
                        
                        If List1.ListCount = 0 Then
                        Slider1.SetFocus
                        Image4.Picture = LoadPicture(App.Path + "\Pix\stop_off.jpg")
                        Image2.Picture = LoadPicture(App.Path + "\Pix\play_off.jpg")
                        Image3.Picture = LoadPicture(App.Path + "\Pix\pausa_off.jpg")
                        MediaPlayer1.Stop
                        MediaPlayer1.FileName = ""
                        Playing = False
                        Paro = True
                        Timer6.Enabled = False
                        Slider1.Value = 0
                        End If
                
                Label1.Caption = ""
            End If
    
Else
MsgBox "No hay ningún elemento en la lista de reproducción", vbOKOnly + vbExclamation, "Error"
Slider1.SetFocus
End If



If List1.ListCount > 0 Then
            
            If Lugar < -1 Then
                Lugar = -1
            End If
            
List1.ListIndex = Lugar
    
    If List1.ListIndex = -1 Then
        List1.Selected(0) = True
    End If
                
                TxT = List1.List(List1.ListIndex)
                TenerNombre
End If

End If
End Sub


Private Sub Image9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    If Image9.Picture = SecondImage Then
        Timer3.Enabled = False
        Image9.Picture = LoadPicture(App.Path + "\Pix\loop_on.jpg")
        Image9.ToolTipText = "Desactivar repetición"
        Replay = True
        Namecito = ""
    Else
        Image9.Picture = LoadPicture(App.Path + "\Pix\loop_off.jpg")
        Image9.ToolTipText = "Activar repetición"
        Replay = False
        Timer3.Enabled = True
        Namecito = ""
    End If
End If
End Sub






Private Sub List1_DblClick()
If Pausa = True Then
MediaPlayer1.Stop
Pausa = False
End If

MUPlay

End Sub


Private Sub List1_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyDown Or vbKeyUp Then
    KDown
End If
     
    If KeyCode = vbKeyReturn Then
        Call List1_DblClick
    End If
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    KDown
End If
End Sub

Private Sub mabout_Click()
Form3.Show
SetWindowPos Form3.hWnd, HWND_TOPMOST, Form3.Left / 15, Form3.Top / 15, Form3.Width / 15, Form3.Height / 15, SWP_NOACTIVATE Or SWP_SHOWWINDOW
End Sub
Private Sub mdirectory_Click()
Form2.Show
SetWindowPos Form2.hWnd, HWND_TOPMOST, Form2.Left / 15, Form2.Top / 15, Form2.Width / 15, Form2.Height / 15, SWP_NOACTIVATE Or SWP_SHOWWINDOW
End Sub

Private Sub mfile_Click()
Dim Libre As Integer
Dim txtdir As String

On Error Resume Next

Libre = FreeFile
Open (App.Path + "\dir.ydr") For Input As Libre
Line Input #Libre, txtdir
Close Libre

CD1.InitDir = txtdir


CD1.ShowOpen
If Not CD1.FileName = "" Then
List1.AddItem CD1.FileName
End If

'*****************************************

If List1.ListCount = 1 Then
        If CD1.FileName <> "" Then
            List1.Selected(0) = True
            TxT = List1.Text
            TenerNombre
        End If
End If
End Sub

Private Sub mminimize_Click()
Form1.WindowState = 1
End Sub

Private Sub montop_Click()

If montop.Checked = False Then
        montop.Checked = True
Else
        montop.Checked = False
End If

'**********************************

If montop.Checked Then
    SetWindowPos Form1.hWnd, HWND_TOPMOST, Form1.Left / 15, Form1.Top / 15, Form1.Width / 15, Form1.Height / 15, SWP_NOACTIVATE Or SWP_SHOWWINDOW
Else
    SetWindowPos Form1.hWnd, HWND_NOTOPMOST, Form1.Left / 15, Form1.Top / 15, Form1.Width / 15, Form1.Height / 15, SWP_NOACTIVATE Or SWP_SHOWWINDOW
End If
End Sub

Private Sub msearch_Click()
Form4.Show
SetWindowPos Form4.hWnd, HWND_TOPMOST, Form4.Left / 15, Form4.Top / 15, Form4.Width / 15, Form4.Height / 15, SWP_NOACTIVATE Or SWP_SHOWWINDOW
End Sub

Private Sub Slider1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Slider1.MousePointer = ccCross
End Sub

Private Sub Slider1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Slider1.MousePointer = ccDefault
End Sub

Private Sub Slider1_Scroll()
MediaPlayer1.CurrentPosition = Slider1.Value
End Sub

Private Sub Slider2_Scroll()
Dim Vol As Long

Vol = Slider2.Value - 2500
MediaPlayer1.Volume = Vol

End Sub


Private Sub Slider3_Scroll()
MediaPlayer1.Balance = Slider3.Value
 
    If Slider3.Value > 500 Then
        Label3.ForeColor = RGB(0, 0, (20 + ((Slider3.Value / 100) * 5)))
        Label3.Caption = "DERECHO"
    ElseIf Slider3.Value < -500 Then
        Label3.ForeColor = RGB((20 - ((Slider3.Value / 100) * 5)), 0, 0)
        Label3.Caption = "IZQUIERDO"
    ElseIf Slider3.Value > -500 And Slider3.Value < 500 Then
        Label3.ForeColor = &HFFFF&
        Label3.Caption = "CENTRO"
    End If
End Sub

Private Sub TiempO_Click()
    Slider1.SetFocus
End Sub

Private Sub Timer1_Timer()

Slider1.Value = MediaPlayer1.CurrentPosition
Slider3.Value = MediaPlayer1.Balance
        
        If Slider2.Value <= 0 Then
        MediaPlayer1.Mute = True
            Else
        MediaPlayer1.Mute = False
        End If

If Slider1.Value = Slider1.Max Then 'Empieza slider1.value = slider1.max
SeDetuvo = True
Slider1.Value = 0
Image2.Picture = LoadPicture(App.Path + "\Pix\play_off.jpg")
Paro = True
Timer6.Enabled = False
Playing = False
Pausa = False
Timer1.Enabled = False
        
        
    If Continuar = True Then 'empieza continuar=true
            If Shuffle = False Then 'Empieza shuffle=false
                
                            If List1.ListCount > 0 Then
                                    If Not List1.ListIndex = (List1.ListCount - 1) Then
                                    List1.Selected(List1.ListIndex + 1) = True
                                    Call List1_DblClick
                                        Else
                                                If Replay = True Then
                                                List1.Selected(0) = True
                                                Call List1_DblClick
                                                    Else
                                                Exit Sub
                                                End If
                                    End If
                            End If
            End If 'Acaba shuffle=false

Else

                    If Replay = True Then
                    List1.Selected(List1.ListIndex) = True
                    Call List1_DblClick
                    End If
            


End If 'Acaba continuar=true

If SeDetuvo = True Then
SigPlay = True
alAzar
End If

End If 'acaba slider1.value = slider1.max

End Sub

Private Sub MUPlay(Optional botton As Integer, Optional X As Single, Optional Y As Single)
Dim Recuerdo As Integer


On Error GoTo heaven

If Pausa = False Then
MediaPlayer1.FileName = List1.Text
MediaPlayer1.Play
Timer6.Enabled = True
Slider1.Max = MediaPlayer1.Duration
Timer1.Enabled = True
Playing = True
Paro = False
TxT = MediaPlayer1.FileName
TenerNombre
Recuerdo = InStr(1, Namecito, MediaPlayer1.FileName, vbTextCompare)

    If Recuerdo = 0 Then
    Namecito = Namecito & MediaPlayer1.FileName & "|"
    End If

Image2.Picture = LoadPicture(App.Path + "\Pix\play_on.jpg")
Timer5.Enabled = True
Else
        If Paro = True Then
        MediaPlayer1.FileName = List1.Text
        Slider1.Max = MediaPlayer1.Duration
        MediaPlayer1.Play
        Timer6.Enabled = True
        TxT = MediaPlayer1.FileName
        TenerNombre
        Recuerdo = InStr(1, Namecito, MediaPlayer1.FileName, vbTextCompare)

            If Recuerdo = 0 Then
        Namecito = Namecito & MediaPlayer1.FileName & "|"
            End If
        
        Timer1.Enabled = True
        Playing = True
        Paro = False
        Timer5.Enabled = True
        End If

Image3.Picture = LoadPicture(App.Path + "\Pix\pausa_off.jpg")
MediaPlayer1.Play
Timer6.Enabled = True
TxT = MediaPlayer1.FileName
TenerNombre
Recuerdo = InStr(1, Namecito, MediaPlayer1.FileName, vbTextCompare)

        If Recuerdo = 0 Then
    Namecito = Namecito & MediaPlayer1.FileName & "|"
        End If
    
Timer1.Enabled = True
Playing = True
Paro = False
Timer5.Enabled = True
End If

If Volumen1 = True Then
    For i = 0 To Slider2.Max
        Slider2.Value = i
    Next i
MediaPlayer1.Volume = Vol
End If

Volumen1 = False

heaven:

        
                If MediaPlayer1.FileName = "" Then
                            If Not List1.Text = "" Then
                            MsgBox "La tarjeta de sonido esta siendo utilizada por otra aplicación!," & Chr$(13) & "o la direccion del archivo no es valida!", vbOKOnly + vbExclamation, "Error"
                            Image2.Picture = LoadPicture(App.Path + "\Pix\play_off.jpg")
                            Exit Sub
                            End If
                MsgBox "  No hay un archivo para reproducir!, " & Chr$(13) & "o la direccion del archivo no es valida!", vbOKOnly + vbExclamation, "Error"
                Image2.Picture = LoadPicture(App.Path + "\Pix\play_off.jpg")
                End If
                
                
                
End Sub

Private Sub Timer2_Timer()
FirstImage = Image6.Picture
End Sub

Sub TenerNombre()
Dim TxtR As String

If TxT = "" Then Exit Sub

TxtR = StrReverse(TxT)
Eslabon = InStr(TxtR, "\")
Label = Left(TxtR, (Eslabon - 1))
Label = StrReverse(Label)
Label = Left(Label, (Len(Label) - 4))
Label = UCase(Label)
Label1.Caption = Label

End Sub

Private Sub Timer3_Timer()
SecondImage = Image9.Picture
End Sub
Private Sub Timer4_Timer()
ThirdImage = Image12.Picture
End Sub

Private Sub Timer5_Timer()

If Playing = True Then
        TxT = MediaPlayer1.FileName
        TenerNombre
        End If
      
End Sub

Private Sub KDown()
    
    Timer5.Enabled = False
    TxT = List1.List(List1.ListIndex)
    TenerNombre
    Timer5.Enabled = True

End Sub

Private Sub SaveList()
Dim CadaUno As String

On Error Resume Next

If Salida = True Then
    ListName = (App.Path + "\DefaultList\Dlist.ply")
End If


    Open ListName For Output As #1

        For n = 0 To (List1.ListCount - 1)
          CadaUno = List1.List(n)
            Print #1, CadaUno
        Next n
        
    Close #1
End Sub

Private Sub LoadList()
Dim EveryOne As String
    
On Error GoTo adios

If Entrada = True Then
    ListName = (App.Path + "\DefaultList\Dlist.ply")
End If
    
    Open ListName For Input As #1
        
        Do While Not EOF(1)
            Line Input #1, EveryOne
            
            If EveryOne <> "" Then
                List1.AddItem EveryOne
            End If
        Loop
        
    Close #1
    
    Entrada = False
    
    Exit Sub
                        

                   
adios:
    
        If Entrada = False Then
            MsgBox "El archivo no fue hallado!", vbOKOnly + vbExclamation, "Error"
        Else
            Entrada = False
            Exit Sub
        End If
        


End Sub

Private Sub ElTiempo()
Dim Sec As Integer
Dim Secs As Integer
Dim Min As Single
Dim Ecsit As Boolean

                
        
Secs = MediaPlayer1.CurrentPosition
Min = Secs / 60
Min = Int(Min)
Sec = Secs - (Min * 60)

    If Playing = False Then
        If Pausa = False Then
If Sec = "-1" Then Sec = "00"
If Not Min = "00" Then Min = "00"
        End If
    End If
    
    If Min > 9 Then
Ecsit = True
    End If
            
            
            
            If Sec < 10 Then
                If Ecsit = False Then
                    TiempO.Text = "0" & Min & ":0" & Sec
                Else
                    TiempO.Text = Min & ":0" & Sec
                End If
            
        Else
                If Ecsit = False Then
                    TiempO.Text = "0" & Min & ":" & Sec
                Else
                    TiempO.Text = Min & ":" & Sec
                End If
            
            End If
                  
                  
                  If Playing = False Then
                    If Pausa = False Then
                        TiempO.Text = "00:00"
                    End If
                  End If
        
End Sub

Private Sub Timer6_Timer()
ElTiempo
End Sub

Sub CargarImagenes()

Image1.Picture = LoadPicture(App.Path + "\Pix\abrir_off.jpg")
Image2.Picture = LoadPicture(App.Path + "\Pix\play_off.jpg")
Image3.Picture = LoadPicture(App.Path + "\Pix\pausa_off.jpg")
Image4.Picture = LoadPicture(App.Path + "\Pix\stop_off.jpg")
Image5.Picture = LoadPicture(App.Path + "\Pix\salir_off.jpg")
Image6.Picture = LoadPicture(App.Path + "\Pix\cont_on.jpg")
Image7.Picture = LoadPicture(App.Path + "\Pix\clear_off.jpg")
Image8.Picture = LoadPicture(App.Path + "\Pix\rem_off.jpg")
Image9.Picture = LoadPicture(App.Path + "\Pix\loop_off.jpg")
Image12.Picture = LoadPicture(App.Path + "\Pix\shuffle_off.jpg")
Image13.Picture = LoadPicture(App.Path + "\Pix\save_off.jpg")
Image14.Picture = LoadPicture(App.Path + "\Pix\load_off.jpg")

End Sub

Sub Seguro()
Dim EstasSeguro As String

CD2.FileName = ""
CD2.ShowSave
ListName = CD2.FileName

If Not ListName <> "" Then Exit Sub

If Dir(ListName, vbHidden) = "" Then
SaveList
Else
EstasSeguro = MsgBox("El nombre de archivo ya existe desea reemplazarlo?", vbYesNo + vbExclamation, "El archivo ya existe")
        
        If EstasSeguro = vbYes Then
        SaveList
        Else
        Seguro
        End If

End If

End Sub

Sub alAzar()
Dim Ancho As Integer
Dim Azar As String


      
                If Shuffle = True Then
                GoTo SueRTe
                    Else
                Exit Sub
                End If
                
SueRTe:
                    'Aqui coloco lo que me va a retener la reproduccion al azar
                    Randomize
                    Ancho = ((-1) * Rnd) + 1 + (List1.ListCount * Rnd) - 1
                    Azar = Int(Ancho)
                    If Azar > (List1.ListCount - 1) Or Azar < 0 Then GoTo SueRTe
                    
                            If List1.ListCount > 1 Then
                            List1.Selected(Azar) = True
                            End If
                               
                                    If Replay = False Then 'empieza replay=false
                                    
                                            If Namecito <> "" Then
                                                     
                                                    Nameci = Left(Namecito, (Len(Namecito) - 1))
                                                        Usados = Split(Nameci, "|")
                                                    
                                                        For i = LBound(Usados) To UBound(Usados)
                                                                            
                                                                        
                                                            If List1.ListCount = (UBound(Usados) + 1) Then
                                                            TiempO.Text = "00:00"
                                                            Timer5.Enabled = False
                                                            TxT = List1.Text
                                                            TenerNombre
                                                            Exit Sub
                                                            End If
                                                                        
                                                                            
                                                            If List1.Text = Usados(i) Then GoTo SueRTe
                                                                                                                                                            
                                                        Next i
                                            
                                                If SigPlay = False Then GoTo la_Exit

                                            
                                            Call List1_DblClick
                                            
                                        Else
                                            
                                            Call List1_DblClick
                                            Exit Sub
                                            
                                            End If
                                Else
                                    
                                    Call List1_DblClick
                                    
                                    End If 'termina replay=false

If Playing = True Then
SeDetuvo = False
End If

la_Exit:
End Sub

Sub irSiguiente()

If Not List1.ListCount > 0 Then Exit Sub 'si no hay nada en la lista de reproduccion

'si la reproduccion al azar esta apagada:
If Shuffle = False Then
    'si la posicion se pasa del final de la lista entonces vuelve al principio(0)
    If List1.ListIndex = (List1.ListCount - 1) Then
    List1.Selected(0) = True
    Else
    'si no se pasa entonces avanza al siguiente
    List1.Selected(List1.ListIndex + 1) = True
    End If
'llamo al procedimiento que filtra el nombre de la cancion
TxT = List1.Text
TenerNombre
        'si esta reproduciendo inmediatamente se pasa a la siguiente cancion y se reproduce
        If Playing = True Then
        Call List1_DblClick
        End If
Else 'si la reproduccion al azar esta prendida:
    If Playing = True Then
    SigPlay = True
    alAzar
    End If
SigPlay = False
'alAzar
End If
End Sub


