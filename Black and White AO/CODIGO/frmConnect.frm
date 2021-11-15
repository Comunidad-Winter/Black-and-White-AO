VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmConnect 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Argentum Online"
   ClientHeight    =   8565
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   FillColor       =   &H00000040&
   Icon            =   "frmConnect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmConnect.frx":000C
   ScaleHeight     =   571
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer Cmd 
      Interval        =   1000
      Left            =   8940
      Top             =   180
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   150
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox PasswordTxt 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   630
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   4980
      Width           =   3615
   End
   Begin VB.TextBox NameTxt 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   690
      TabIndex        =   4
      Top             =   3570
      Width           =   3615
   End
   Begin VB.ListBox lst_servers 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   3180
      ItemData        =   "frmConnect.frx":65C97
      Left            =   -3090
      List            =   "frmConnect.frx":65C9E
      TabIndex        =   3
      Top             =   9060
      Visible         =   0   'False
      Width           =   3045
   End
   Begin VB.TextBox PortTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   1260
      TabIndex        =   0
      Text            =   "7666"
      Top             =   8670
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.TextBox IPTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   3390
      TabIndex        =   2
      Text            =   "localhost"
      Top             =   8640
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label lblestado 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Comprobando..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   405
      Left            =   9960
      TabIndex        =   6
      Top             =   630
      Width           =   1695
   End
   Begin VB.Image imgServEspana 
      Height          =   195
      Left            =   -870
      MousePointer    =   99  'Custom
      Top             =   9060
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.Image imgServArgentina 
      Height          =   495
      Left            =   -1890
      MousePointer    =   99  'Custom
      Top             =   9060
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.Image imgGetPass 
      Height          =   255
      Left            =   -180
      MousePointer    =   99  'Custom
      Top             =   9060
      Width           =   3795
   End
   Begin VB.Label version 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Image Image1 
      Height          =   1125
      Index           =   0
      Left            =   7680
      MousePointer    =   99  'Custom
      Top             =   7080
      Width           =   3630
   End
   Begin VB.Image Image1 
      Height          =   1095
      Index           =   1
      Left            =   570
      MousePointer    =   99  'Custom
      Top             =   7080
      Width           =   3615
   End
   Begin VB.Image Image1 
      Height          =   150
      Index           =   2
      Left            =   -720
      MousePointer    =   99  'Custom
      Top             =   9060
      Width           =   2490
   End
   Begin VB.Image FONDO 
      Height          =   360
      Left            =   11100
      Top             =   9060
      Visible         =   0   'False
      Width           =   2820
   End
End
Attribute VB_Name = "frmConnect"
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
'
'Matías Fernando Pequeño
'matux@fibertel.com.ar
'www.noland-studios.com.ar
'Acoyte 678 Piso 17 Dto B
'Capital Federal, Buenos Aires - Republica Argentina
'Código Postal 1405

Option Explicit



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
        Call SaveGameini
        frmConnect.MousePointer = 1
        frmMain.MousePointer = 1
        prgRun = False
        Call UnloadAllForms
End If
End Sub

Private Sub Form_Load()
If Winsock1.State <> sckClosed Then
Winsock1.Close
End If
Winsock1.Connect "127.0.0.1", "7668"
    '[CODE 002]:MatuX
    EngineRun = False
    '[END]
    
 Dim j
 For Each j In Image1()
    j.Tag = "0"
 Next
 frmConnect.Picture = LoadPicture(App.path & "\Graficos\Conectar.jpg")
 '[CODE]:MatuX
 '
 '  El código para mostrar la versión se genera acá para
 ' evitar que por X razones luego desaparezca, como suele
 ' pasar a veces :)
    version.Caption = "v" & App.Major & "." & App.Minor & " Build: " & App.Revision
 '[END]'

End Sub



Private Sub Image1_Click(Index As Integer)


CurServer = 0
IPdelServidor = "127.0.0."
PuertoDelServidor = "7666"


Call Audio.PlayWave(SND_CLICK)

Select Case Index
    Case 0
        
       EstadoLogin = CrearAccount
#If UsarWrench = 1 Then
       If frmMain.Socket1.Connected Then
            frmMain.Socket1.Disconnect
            frmMain.Socket1.Cleanup
            DoEvents
        End If
        frmMain.Socket1.HostAddress = CurServerIp
        frmMain.Socket1.RemotePort = CurServerPort
        frmMain.Socket1.Connect
#Else
        If frmMain.Winsock1.State <> sckClosed Then
            frmMain.Winsock1.Close
            DoEvents
        End If
        frmMain.Winsock1.Connect CurServerIp, CurServerPort
#End If

        
    Case 1
    
#If UsarWrench = 1 Then
        If frmMain.Socket1.Connected Then
        frmMain.Socket1.Disconnect
        frmMain.Socket1.Cleanup
        DoEvents
        End If
#Else
        If frmMain.Winsock1.State <> sckClosed Then _
            frmMain.Winsock1.Close
#End If
      '  If frmConnect.MousePointer = 99 Then
      '      Exit Sub
     '   End If
        
        
        'update user info
        UserName = NameTxt.Text
        NameRelogueo = UserName
        Dim aux As String
        aux = PasswordTxt.Text
        PassRelogueo = aux
#If SeguridadAlkon Then
        UserPassword = md5.GetMD5String(aux)
        Call md5.MD5Reset
#Else
        UserPassword = aux
#End If
        If CheckUserData(False) = True Then
            EstadoLogin = loginaccount
            Me.MousePointer = 99
#If UsarWrench = 1 Then
            frmMain.Socket1.HostAddress = CurServerIp
            frmMain.Socket1.RemotePort = CurServerPort
            frmMain.Socket1.Connect
#Else
            'If frmMain.Winsock1.State <> sckClosed Then _
               ' frmMain.Winsock1.Close
            frmMain.Winsock1.Connect CurServerIp, CurServerPort
#End If
        End If
        

End Select
Exit Sub


End Sub

Private Sub PasswordTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call Image1_Click(1)
    End If
End Sub

Private Sub Winsock1_Connect()
lblestado.ForeColor = vbGreen
lblestado.Caption = "Online"
End Sub
 
Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
lblestado.ForeColor = vbRed
lblestado.Caption = "Offline"
End Sub
