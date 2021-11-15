VERSION 5.00
Begin VB.Form frmOpciones 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4665
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOpciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      BackColor       =   &H00000000&
      Caption         =   "Pantalla"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1590
      TabIndex        =   18
      Top             =   4800
      Width           =   1455
      Begin VB.CheckBox CheckFps 
         BackColor       =   &H00000000&
         Caption         =   "Ver Fps"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   270
         TabIndex        =   19
         Top             =   210
         Width           =   945
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "Seleccion de Interfaz"
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   720
      TabIndex        =   16
      Top             =   5520
      Width           =   3195
      Begin VB.ListBox Interfaces 
         Height          =   1230
         ItemData        =   "frmOpciones.frx":0152
         Left            =   420
         List            =   "frmOpciones.frx":0168
         TabIndex        =   17
         Top             =   390
         Width           =   2355
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Configuracion de Sonido"
      ForeColor       =   &H00FFFFFF&
      Height          =   1035
      Left            =   540
      TabIndex        =   11
      Top             =   3690
      Width           =   3525
      Begin VB.HScrollBar Slider2 
         Height          =   255
         Left            =   1140
         Max             =   100
         TabIndex        =   13
         Top             =   660
         Value           =   50
         Width           =   2085
      End
      Begin VB.HScrollBar Slider1 
         Height          =   255
         Left            =   1140
         Max             =   100
         TabIndex        =   12
         Top             =   330
         Value           =   50
         Width           =   2085
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Musica"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   420
         TabIndex        =   15
         Top             =   660
         Width           =   645
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Sonido"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   420
         TabIndex        =   14
         Top             =   360
         Width           =   645
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Configurar Teclas"
      Height          =   345
      Left            =   960
      TabIndex        =   10
      Top             =   870
      Width           =   2790
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Luz en el Mouse"
      Height          =   345
      Left            =   960
      TabIndex        =   9
      Top             =   1350
      Width           =   2790
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Diálogos de clan"
      ForeColor       =   &H00FFFFFF&
      Height          =   750
      Left            =   240
      TabIndex        =   4
      Top             =   2790
      Width           =   4230
      Begin VB.TextBox txtCantMensajes 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2925
         MaxLength       =   1
         TabIndex        =   7
         Text            =   "5"
         Top             =   315
         Width           =   450
      End
      Begin VB.OptionButton optPantalla 
         BackColor       =   &H00000000&
         Caption         =   "En pantalla,"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   1770
         TabIndex        =   6
         Top             =   315
         Value           =   -1  'True
         Width           =   1560
      End
      Begin VB.OptionButton optConsola 
         BackColor       =   &H00000000&
         Caption         =   "En consola"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   105
         TabIndex        =   5
         Top             =   315
         Width           =   1560
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "mensajes"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   3480
         TabIndex        =   8
         Top             =   345
         Width           =   750
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cerrar"
      Height          =   345
      Left            =   870
      MouseIcon       =   "frmOpciones.frx":01D1
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   7770
      Width           =   2790
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sonidos Activados"
      Height          =   345
      Index           =   1
      Left            =   960
      MouseIcon       =   "frmOpciones.frx":0323
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   2340
      Width           =   2790
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Musica Activada"
      Height          =   345
      Index           =   0
      Left            =   960
      MouseIcon       =   "frmOpciones.frx":0475
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   1830
      Width           =   2790
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Opciones"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.9.0.9
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit



Private Sub Command1_Click(index As Integer)

Call Audio.PlayWave(SND_CLICK)

Select Case index
    Case 0
        If Musica Then
            Musica = False
            Command1(0).Caption = "Musica Desactivada"
            Audio.StopMidi
        Else
            Musica = True
            Command1(0).Caption = "Musica Activada"
            Call Audio.PlayMIDI(CStr(currentMidi) & ".mid")
        End If
    Case 1
    
        If Sound Then
            Sound = False
            Command1(1).Caption = "Sonidos Desactivados"
            Call Audio.StopWave
            RainBufferIndex = 0
            frmMain.IsPlaying = PlayLoop.plNone
        Else
            Sound = True
            Command1(1).Caption = "Sonidos Activados"
        End If
End Select
End Sub

Private Sub Command2_Click()
Me.Visible = False
End Sub

Private Sub Command3_Click()
LuzMouse = Not LuzMouse
    If Not LuzMouse Then
        Engine.Light_Remove (Engine.Light_Find(20))
    Else
        Engine.Light_Create UserPos.X + frmMain.MouseX \ 32 - frmMain.renderer.ScaleWidth \ 64, UserPos.y + frmMain.MouseY / 32 - frmMain.renderer.ScaleHeight \ 64, D3DColorXRGB(255, 255, 255), 2, 20
    End If
End Sub

Private Sub Command4_Click()
Call frmCustomKeys.Show(vbModeless, frmMain)
End Sub

Private Sub Form_Load()
    If Musica Then
        Command1(0).Caption = "Musica Activada"
    Else
        Command1(0).Caption = "Musica Desactivada"
    End If
    
    If Sound Then
        Command1(1).Caption = "Sonidos Activados"
    Else
        Command1(1).Caption = "Sonidos Desactivados"
    End If
End Sub

 Private Sub Interfaces_Click()
Select Case Interfaces
Case "Clasica"
NumSkin = 1
frmMain.Picture = LoadPicture(App.path & "\Graficos\Todo.jpg") ' Nombre de la interfaz del main.
frmMain.InvEqu.Picture = LoadPicture(App.path & "\Graficos\Centronuevohechizos.jpg") 'Nombre de la interfaz del inventario
frmMain.InvEqu.Picture = LoadPicture(App.path & "\Graficos\Centronuevoinventario.jpg") ' Nombre de la interfaz del hechizo
Case "Bosque maldito"
NumSkin = 2
frmMain.Picture = LoadPicture(App.path & "\Graficos\Todo1.jpg")
frmMain.InvEqu.Picture = LoadPicture(App.path & "\Graficos\Centronuevohechizos1.jpg")
frmMain.InvEqu.Picture = LoadPicture(App.path & "\Graficos\Centronuevoinventario1.jpg")
Case "Dragones en el mar"
NumSkin = 3
frmMain.Picture = LoadPicture(App.path & "\Graficos\Todo2.jpg")
frmMain.InvEqu.Picture = LoadPicture(App.path & "\Graficos\Centronuevohechizos2.jpg")
frmMain.InvEqu.Picture = LoadPicture(App.path & "\Graficos\Centronuevoinventario2.jpg")
Case "Guerrero malvado"
NumSkin = 4
frmMain.Picture = LoadPicture(App.path & "\Graficos\Todo3.jpg")
frmMain.InvEqu.Picture = LoadPicture(App.path & "\Graficos\Centronuevohechizos3.jpg")
frmMain.InvEqu.Picture = LoadPicture(App.path & "\Graficos\Centronuevoinventario3.jpg")
Case "Pelea de dragones"
NumSkin = 5
frmMain.Picture = LoadPicture(App.path & "\Graficos\Todo4.jpg")
frmMain.InvEqu.Picture = LoadPicture(App.path & "\Graficos\Centronuevohechizos4.jpg")
frmMain.InvEqu.Picture = LoadPicture(App.path & "\Graficos\Centronuevoinventario4.jpg")
Case "Dragon legendario"
NumSkin = 6
frmMain.Picture = LoadPicture(App.path & "\Graficos\Todo5.jpg")
frmMain.InvEqu.Picture = LoadPicture(App.path & "\Graficos\Centronuevohechizos5.jpg")
frmMain.InvEqu.Picture = LoadPicture(App.path & "\Graficos\Centronuevoinventario5.jpg")
End Select
 
End Sub

Private Sub optConsola_Click()
    DialogosClanes.Activo = False
End Sub

Private Sub optPantalla_Click()
    DialogosClanes.Activo = True
End Sub

Private Sub Slider1_Change()
Audio.SoundVolume = Slider1.value
End Sub

Private Sub Slider2_Change()
Audio.MusicVolume = Slider2.value
End Sub

Private Sub txtCantMensajes_LostFocus()
    txtCantMensajes.Text = Trim$(txtCantMensajes.Text)
    If IsNumeric(txtCantMensajes.Text) Then
        DialogosClanes.CantidadDialogos = Trim$(txtCantMensajes.Text)
    Else
        txtCantMensajes.Text = 5
    End If
End Sub
