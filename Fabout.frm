VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H0000FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acercade Yova's MPlayer"
   ClientHeight    =   3930
   ClientLeft      =   3210
   ClientTop       =   2370
   ClientWidth     =   5760
   ControlBox      =   0   'False
   Icon            =   "Fabout.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   262
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   384
   Begin VB.Timer Bajadita 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   600
      Top             =   3120
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H8000000B&
      Caption         =   "Acept&ar"
      Default         =   -1  'True
      Height          =   555
      Left            =   2280
      MaskColor       =   &H00FFC0C0&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3120
      Width           =   1185
   End
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1020
      Left            =   2370
      ScaleHeight     =   695.31
      ScaleMode       =   0  'User
      ScaleWidth      =   695.31
      TabIndex        =   1
      Top             =   1920
      Width           =   1020
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Index           =   5
      X1              =   8
      X2              =   376
      Y1              =   8
      Y2              =   8
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      Index           =   4
      X1              =   376
      X2              =   376
      Y1              =   8
      Y2              =   256
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      Index           =   3
      X1              =   8
      X2              =   376
      Y1              =   256
      Y2              =   256
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      Index           =   2
      X1              =   8
      X2              =   8
      Y1              =   8
      Y2              =   256
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "yovas@usa.net"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   938
      TabIndex        =   4
      Top             =   1200
      Width           =   3885
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "5.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   435
      Left            =   2633
      TabIndex        =   3
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Yova's MPlayer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   480
      Left            =   938
      TabIndex        =   2
      Top             =   120
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Index           =   1
      X1              =   8
      X2              =   376
      Y1              =   72
      Y2              =   72
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      Index           =   0
      X1              =   8
      X2              =   376
      Y1              =   120
      Y2              =   120
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Bajar As Long

Private Sub cmdOK_Click()
Unload Form3
End Sub


Private Sub Form_Load()
Dim Rgn1 As Long
Dim Rgn2 As Long
Dim Rc As Long

    Rgn1 = CreateRectRgn(0, 0, 0, 0)
    Rc = SetWindowRgn(Me.hWnd, Rgn1, True)

Bajar = 0
Form3.Bajadita.Enabled = True

picIcon.Picture = LoadPicture(App.Path + "\Pix\icono.ico")

End Sub


Private Sub Bajadita_Timer()
Dim Rgn1 As Long
Dim Rgn2 As Long
Dim Rc As Long

    Rgn1 = CreateRectRgn(5, 0, 384, (0 + Bajar))
    Rc = SetWindowRgn(Me.hWnd, Rgn1, True)
    Bajar = Bajar + 4
        
        If Bajar >= 286 Then
        Form3.Bajadita.Enabled = False
        Exit Sub
        End If
    
End Sub
