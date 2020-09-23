VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H0000FFFF&
   Caption         =   "Abrir Directorio"
   ClientHeight    =   3720
   ClientLeft      =   4740
   ClientTop       =   1995
   ClientWidth     =   2775
   Icon            =   "Directorios.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   248
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   185
   Begin VB.CheckBox Check1 
      BackColor       =   &H0000FFFF&
      Caption         =   "INCLUIR SUBDIRECTORIOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   180
      TabIndex        =   4
      Top             =   3120
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   120
      ScaleHeight     =   3015
      ScaleWidth      =   2535
      TabIndex        =   0
      Top             =   120
      Width           =   2535
      Begin VB.DirListBox Dir1 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FFFF&
         Height          =   1665
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   2055
      End
      Begin VB.FileListBox File1 
         Height          =   1065
         Left            =   600
         Pattern         =   "*.mp3"
         TabIndex        =   2
         Top             =   840
         Width           =   1215
      End
      Begin VB.DriveListBox Drive1 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FFFF&
         Height          =   315
         Left            =   240
         TabIndex        =   1
         Top             =   0
         Width           =   2055
      End
      Begin VB.Image Image2 
         Height          =   450
         Left            =   525
         Top             =   2520
         Width           =   1470
      End
      Begin VB.Image Image1 
         Height          =   450
         Left            =   810
         Top             =   2040
         Width           =   915
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public txtdir As String
Public Libre As Integer


Private Sub Check1_Click()
Static Conta_Pues As Integer


    If Conta_Pues > 0 Then Exit Sub

If Check1.Value = 1 Then
        
        MsgBox "Debido a lo complejo de este algoritmo, pueden haber errores en el alcance de los subdirectorios!", , "Atención"
        Conta_Pues = Conta_Pues + 1

End If

End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Image1.Picture = LoadPicture(App.Path + "\Pix\ok_on.jpg")
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Song As String
Dim El_Dir As String
Dim El_Dir2 As String
Dim El_Dir3 As String


        
        If Button = 1 Then
        
Image1.Picture = LoadPicture(App.Path + "\Pix\ok_off.jpg")
        
        Form5.Show ' se muestra la ventana que contiene la barra de progreso
        Form2.Visible = False 'desaparece el formulario para que no haga esfuerzos banos en mostrar la operación
        Form1.Refresh
        
SetWindowPos Form5.hWnd, HWND_TOPMOST, Form5.Left / 15, Form5.Top / 15, Form5.Width / 15, Form5.Height / 15, SWP_NOACTIVATE Or SWP_SHOWWINDOW

File1.Path = Dir1.Path
       
'Se guarda el path del directorio para usarlo despues

Open (App.Path & "\" + "dir.ydr") For Output As Libre
Print #Libre, Dir1.Path
Close Libre
       
'Si la opcion de buscar en subdirectorios esta activada se recurre al procedimiento Buscame
       
If Check1.Value = 1 Then

Form5.ProgressBar1.Max = 300 'el nuevo valor maximo de la barra de progreso

File1.Path = Dir1.Path
El_Dir = Dir1.Path

        For i = 0 To File1.ListCount - 1
            If Len(Dir1.Path) > 3 Then
                Song = Dir1.Path & "\" & File1.List(i)
                Else
                Song = Dir1.Path & File1.List(i)
                End If
        
        Form1.List1.AddItem Song
    
        Next i

        
For n = 0 To (Dir1.ListCount - 1)

Dir1.Path = El_Dir
                
File1.Path = Dir1.List(n)
       
Dir1.Path = File1.Path


                For i = 0 To File1.ListCount - 1
                    If Len(Dir1.Path) > 3 Then
                    Song = Dir1.Path & "\" & File1.List(i)
                    Else
                    Song = Dir1.Path & File1.List(i)
                    End If
        
                Form1.List1.AddItem Song
                Next i

           
     If Dir1.ListCount >= 1 Then
     El_Dir2 = Dir1.Path
     
        For j = 0 To Dir1.ListCount - 1
            Dir1.Path = El_Dir2
            
            File1.Path = Dir1.List(j)
            Dir1.Path = File1.Path
            
            For k = 0 To File1.ListCount - 1
                    If Len(Dir1.Path) > 3 Then
                    Song = Dir1.Path & "\" & File1.List(k)
                    Else
                    Song = Dir1.Path & File1.List(k)
                    End If
        
            Form1.List1.AddItem Song
            Next k
                          
                    If Dir1.ListCount >= 1 Then
                    El_Dir3 = Dir1.Path
                    
                        For z = 0 To Dir1.ListCount - 1
                        Dir1.Path = El_Dir3
                        
                        File1.Path = Dir1.List(z)
                        Dir1.Path = File1.Path
        
                            For r = 0 To File1.ListCount - 1
                                If Len(Dir1.Path) > 3 Then
                                Song = Dir1.Path & "\" & File1.List(r)
                                Else
                                Song = Dir1.Path & File1.List(r)
                                End If
        
                            Form1.List1.AddItem Song
                            Next r
                            
                                        If Dir1.ListCount >= 1 Then
                                     
                                            For o = 0 To Dir1.ListCount - 1
                                            
                        
                                            File1.Path = Dir1.List(o)
                                            
        
                                                    For p = 0 To File1.ListCount - 1
                                                        If Len(Dir1.Path) > 3 Then
                                                        Song = Dir1.Path & "\" & File1.List(p)
                                                        Else
                                                        Song = Dir1.Path & File1.List(p)
                                                        End If
        
                                                    Form1.List1.AddItem Song
                                                    Next p
                    
                                            Next o
                                            
                                        End If
                                    
                        Next z
                    End If
             If Form5.ProgressBar1.Value < 50 Then Form5.ProgressBar1.Value = 50
        
        Next j
     End If
   If Form5.ProgressBar1.Value < 250 Then Form5.ProgressBar1.Value = 250
                
DoEvents

Next n

If Form5.ProgressBar1.Value < 300 Then Form5.ProgressBar1.Value = 300

GoTo En_la_jugada

End If
        
'En otro caso solo carga los archivos del directorio especifico
    
Form5.ProgressBar1.Max = 100 'el maximo de la barra de progreso
    
For i = 0 To (File1.ListCount - 1)
        If Len(Dir1.Path) > 3 Then
        Song = Dir1.Path & "\" & File1.List(i)
        Else
        Song = Dir1.Path & File1.List(i)
        End If
        
     Form1.List1.AddItem Song
   
Form5.ProgressBar1.Value = (i / File1.ListCount) * 100
   
Next i

GoTo En_la_jugada

'Aqui se obtiene el nombre de la cancion 1 de la lista

En_la_jugada:

            If Form1.List1.ListCount > 0 Then
                Form1.List1.Selected(0) = True
                TxT = Form1.List1.Text
                Form1.TenerNombre
            End If
            

        End If
        
        Unload Form2 'cierro la barra de progreso
        
        Unload Form5 'se cierra la ventana de directorios!

                
End Sub




Private Sub Drive1_Change()

On Error Resume Next
Dir1.Path = Drive1.Drive

End Sub


Private Sub Form_Load()

On Error Resume Next

    Rgn1 = CreateRoundRectRgn(10, 25, 183, 200, 150, 30)
    Rgn2 = CreateRoundRectRgn(10, 190, 183, 270, 100, 40)
    Rc = CombineRgn(Rgn1, Rgn1, Rgn2, RGN_OR)
    Rc = SetWindowRgn(Me.hWnd, Rgn1, True)

Libre = FreeFile
Open (App.Path + "\dir.ydr") For Input As Libre
Line Input #Libre, txtdir
Close Libre

Dir1.Path = txtdir

Image1.Picture = LoadPicture(App.Path + "\Pix\ok_off.jpg")
Image2.Picture = LoadPicture(App.Path + "\Pix\can_off.jpg")

End Sub


Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Image2.Picture = LoadPicture(App.Path + "\Pix\can_on.jpg")
End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
    
Image2.Picture = LoadPicture(App.Path + "\Pix\can_off.jpg")
    
Unload Form2
    
    End If
    
End Sub



