VERSION 5.00
Begin VB.Form frmCuent 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11970
   BeginProperty Font 
      Name            =   "Georgia"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmCuent.frx":0000
   MousePointer    =   99  'Custom
   Picture         =   "frmCuent.frx":0CCA
   ScaleHeight     =   9000
   ScaleWidth      =   11970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "BorrarPJ"
      Height          =   315
      Left            =   11790
      TabIndex        =   37
      Top             =   9060
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Deslogear"
      Height          =   375
      Left            =   12030
      TabIndex        =   36
      Top             =   1380
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Conectar"
      Height          =   255
      Left            =   600
      TabIndex        =   35
      Top             =   9060
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Crear PJ"
      Height          =   525
      Left            =   4650
      TabIndex        =   33
      Top             =   9060
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   5
      Left            =   4380
      MouseIcon       =   "frmCuent.frx":456CF
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":46399
      ScaleHeight     =   1215
      ScaleWidth      =   735
      TabIndex        =   32
      Top             =   4830
      Width           =   735
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   6
      Left            =   6570
      MouseIcon       =   "frmCuent.frx":46634
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":472FE
      ScaleHeight     =   1215
      ScaleWidth      =   735
      TabIndex        =   31
      Top             =   4860
      Width           =   735
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   7
      Left            =   8700
      MouseIcon       =   "frmCuent.frx":47599
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":48263
      ScaleHeight     =   1215
      ScaleWidth      =   735
      TabIndex        =   30
      Top             =   4890
      Width           =   735
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   4
      Left            =   2250
      MouseIcon       =   "frmCuent.frx":484FE
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":491C8
      ScaleHeight     =   1215
      ScaleWidth      =   735
      TabIndex        =   4
      Top             =   4860
      Width           =   735
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   3
      Left            =   8730
      MouseIcon       =   "frmCuent.frx":49463
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":4A12D
      ScaleHeight     =   1215
      ScaleWidth      =   735
      TabIndex        =   3
      Top             =   2010
      Width           =   735
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   2
      Left            =   6570
      MouseIcon       =   "frmCuent.frx":4A3C8
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":4B092
      ScaleHeight     =   1215
      ScaleWidth      =   735
      TabIndex        =   2
      Top             =   2010
      Width           =   735
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   1
      Left            =   4380
      MouseIcon       =   "frmCuent.frx":4B32D
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":4BFF7
      ScaleHeight     =   1215
      ScaleWidth      =   735
      TabIndex        =   1
      Top             =   2040
      Width           =   735
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   0
      Left            =   2190
      MouseIcon       =   "frmCuent.frx":4C292
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":4CF5C
      ScaleHeight     =   1215
      ScaleWidth      =   735
      TabIndex        =   0
      Top             =   2010
      Width           =   735
   End
   Begin VB.Image Image4 
      Height          =   1275
      Left            =   8310
      Top             =   7560
      Width           =   3495
   End
   Begin VB.Image Image3 
      Height          =   855
      Left            =   8340
      Top             =   180
      Width           =   2625
   End
   Begin VB.Image Image2 
      Height          =   1275
      Left            =   4350
      Top             =   7530
      Width           =   3525
   End
   Begin VB.Image Image1 
      Height          =   1245
      Left            =   390
      Top             =   7560
      Width           =   3495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre de la Cuenta:"
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   720
      TabIndex        =   34
      Top             =   480
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   8040
      TabIndex        =   29
      Top             =   6840
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   5910
      TabIndex        =   28
      Top             =   6840
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   3720
      TabIndex        =   27
      Top             =   6840
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   8280
      TabIndex        =   26
      Top             =   6630
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   6150
      TabIndex        =   25
      Top             =   6630
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   3960
      TabIndex        =   24
      Top             =   6630
      Width           =   1455
   End
   Begin VB.Label nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   7
      Left            =   8220
      TabIndex        =   23
      Top             =   4500
      Width           =   1815
   End
   Begin VB.Label nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   6
      Left            =   6060
      TabIndex        =   22
      Top             =   4500
      Width           =   1815
   End
   Begin VB.Label nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   5
      Left            =   3840
      TabIndex        =   21
      Top             =   4500
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "PJClick"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3870
      TabIndex        =   20
      Top             =   450
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   1950
      TabIndex        =   19
      Top             =   6630
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   8340
      TabIndex        =   18
      Top             =   3810
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   6090
      TabIndex        =   17
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   3990
      TabIndex        =   16
      Top             =   3810
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   1710
      TabIndex        =   15
      Top             =   6840
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   8100
      TabIndex        =   14
      Top             =   4050
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   5850
      TabIndex        =   13
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   3750
      TabIndex        =   12
      Top             =   4050
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   1680
      TabIndex        =   11
      Top             =   4050
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   1920
      TabIndex        =   10
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   4
      Left            =   1770
      TabIndex        =   9
      Top             =   4530
      Width           =   1815
   End
   Begin VB.Label nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   3
      Left            =   8220
      TabIndex        =   8
      Top             =   1650
      Width           =   1815
   End
   Begin VB.Label nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   2
      Left            =   6030
      TabIndex        =   7
      Top             =   1620
      Width           =   1815
   End
   Begin VB.Label nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   3870
      TabIndex        =   6
      Top             =   1650
      Width           =   1815
   End
   Begin VB.Label nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   0
      Left            =   1770
      TabIndex        =   5
      Top             =   1650
      Width           =   1815
   End
End
Attribute VB_Name = "frmCuent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit



Private Sub Command1_Click()
If PJClickeado = "Nada" Then
MsgBox "Seleccione un pj"
End If
Call Audio.PlayWave(SND_CLICK)
UserName = PJClickeado
SendData ("OOLOGI" & UserName)
Unload Me
End Sub

Private Sub Command4_Click()
frmBorrar.Show , frmCuent
End Sub

Private Sub Form_Load()
Dim i As Integer
Label3.Caption = UserName
End Sub
Private Sub Command2_Click()
frmMain.Socket1.Disconnect
Unload Me
frmConnect.Show
End Sub

Private Sub Command3_Click()
Call Audio.PlayWave(SND_CLICK)

If Nombre(7).Caption <> "Nada" Then
    MsgBox "Tu cuenta ha llegado al máximo de personajes."
    Exit Sub
End If

    EstadoLogin = Dados
    frmCrearPersonaje.Show vbModal
    Me.MousePointer = 11
    
End Sub


Private Sub Image1_Click()
If PJClickeado = "Nada" Then
MsgBox "Seleccione un pj"
End If
Call Audio.PlayWave(SND_CLICK)
UserName = PJClickeado
SendData ("OOLOGI" & UserName)
Unload Me
End Sub

Private Sub Image2_Click()
Call Audio.PlayWave(SND_CLICK)

If Nombre(7).Caption <> "Nada" Then
    MsgBox "Tu cuenta ha llegado al máximo de personajes."
    Exit Sub
End If

    EstadoLogin = Dados
    frmCrearPersonaje.Show vbModal
    Me.MousePointer = 11
    
End Sub

Private Sub Image3_Click()
frmMain.Socket1.Disconnect
Unload Me
frmConnect.Show
End Sub

Private Sub Image4_Click()
frmBorrar.Show , frmCuent
End Sub

Private Sub nombre_dblClick(index As Integer)
If PJClickeado = "Nada" Then Exit Sub
Call Audio.PlayWave(SND_CLICK)
UserName = PJClickeado
SendData ("OOLOGI" & UserName)
Unload Me
End Sub
Private Sub nombre_Click(index As Integer)
PJClickeado = frmCuent.Nombre(index).Caption
End Sub
Private Sub PJ_Click(index As Integer)
PJClickeado = frmCuent.Nombre(index).Caption
End Sub

Private Sub PJ_dblClick(index As Integer)
If PJClickeado = "Nada" Then Exit Sub
Call Audio.PlayWave(SND_CLICK)
UserName = PJClickeado
SendData ("OOLOGI" & UserName)
Unload Me
End Sub


