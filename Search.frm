VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H0000FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar canción"
   ClientHeight    =   1515
   ClientLeft      =   1455
   ClientTop       =   3120
   ClientWidth     =   3690
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   101
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   246
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   4560
      Top             =   1080
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   2400
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   120
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Buscar:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   240
      Top             =   960
      Width           =   1575
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Rayar As Long


Private Sub Form_Load()
Dim Rgn1 As Long
Dim Rc As Long

    Rgn1 = CreateRectRgn(0, 0, 0, 0)
    Rc = SetWindowRgn(Me.hWnd, Rgn1, True)

Rayar = 0
Form4.Timer1.Enabled = True

Image2.Picture = LoadPicture(App.Path + "\Pix\can_off.jpg")
Image3.Picture = LoadPicture(App.Path + "\Pix\searchn_off.jpg")

End Sub



Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Image2.Picture = LoadPicture(App.Path + "\Pix\can_on.jpg")
End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Image2.Picture = LoadPicture(App.Path + "\Pix\can_off.jpg")
Unload Form4
End Sub

Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Image3.Picture = LoadPicture(App.Path + "\Pix\searchn_on.jpg")
End Sub

Private Sub Image3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 1 Then
Image3.Picture = LoadPicture(App.Path + "\Pix\searchn_off.jpg")
BuscarSiguiente
End If

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
'aqui si el usuario decide pulsar ENTER la busqueda se realiza
If KeyCode = vbKeyReturn Then
    If Form4.Text1.Text <> "" Then
    'llama a la funcion que realiza las busquedas
    BuscarSiguiente
    Else
    'si no hay ningun caracter en la textbox despliega un mensaje
    MsgBox "Ingrese el nombre de una canción, o una sus palabras para buscar", vbOKOnly, "Error"
    End If
End If
'si el usuario pulsa ESCAPE cierra el formulario
If KeyCode = vbKeyEscape Then
Unload Form4
End If

End Sub

Sub Timer1_Timer()
Dim Rgn1 As Long
Dim Rc As Long

    Rgn1 = CreateRectRgn(5, 5, (5 + Rayar), 120)
    Rc = SetWindowRgn(Me.hWnd, Rgn1, True)
    Rayar = Rayar + 6
        
        If Rayar >= 246 Then
        Form4.Timer1.Enabled = False
        Exit Sub
        End If
    
End Sub



Private Sub BuscarSiguiente()
Dim SearchChr As Integer
Dim Busqueda As String
Static Puntito As Integer
Static Punto As Integer


If Not Form1.List1.ListCount > 0 Then

MsgBox "La lista de reproduccion esta vacia!", , "Error"
Exit Sub

End If


Busqueda = Form4.Text1.Text
        
        If Not Busqueda <> "" Then Exit Sub
                
                If Form1.List1.ListIndex = 0 Then
                Puntito = 1
                End If
                
        For k = (Punto + Puntito) To Form1.List1.ListCount - 1
            SearchChr = InStr(1, Form1.List1.List(k), Busqueda, vbTextCompare)
                If SearchChr > 0 Then
                    Form1.List1.Selected(k) = True
                    TxT = Form1.List1.List(Form1.List1.ListIndex)
                    Form1.TenerNombre
                    Punto = Form1.List1.ListIndex
                    Puntito = 1
                    Exit For
                 End If
        Next k
        
                If k >= (Form1.List1.ListCount - 1) Then
                MsgBox "No se encontró una palabra que concordará con su busqueda", vbOKOnly, "Busqueda finalizada"
                Punto = 0
                Puntito = 0
                Exit Sub
                End If
                     
End Sub

