Attribute VB_Name = "Duelos"
Option Explicit

   Public Sub ComensarDuelo(ByVal UserIndex As Integer, ByVal tIndex As Integer)
    UserList(UserIndex).flags.EstaDueleando = True
    UserList(UserIndex).flags.Oponente = tIndex
    UserList(tIndex).flags.EstaDueleando = True
    Call WarpUserChar(tIndex, 192, 43, 32)
    UserList(tIndex).flags.Oponente = UserIndex
    Call WarpUserChar(UserIndex, 192, 66, 49)
    Call SendData(ToAll, 0, 0, "||Retos> " & UserList(tIndex).name & " y " & UserList(UserIndex).name & " van a competir en un Reto." & FONTTYPE_TALK)
    End Sub
    Public Sub ResetDuelo(ByVal UserIndex As Integer, ByVal tIndex As Integer)
    UserList(UserIndex).flags.EsperandoDuelo = False
    UserList(UserIndex).flags.Oponente = 0
    UserList(UserIndex).flags.EstaDueleando = False
    Call WarpUserChar(UserIndex, 1, 50, 50)
    Call WarpUserChar(tIndex, 1, 51, 51)
    UserList(tIndex).flags.EsperandoDuelo = False
    UserList(tIndex).flags.Oponente = 0
    UserList(tIndex).flags.EstaDueleando = False
    End Sub
    Public Sub TerminarDuelo(ByVal Ganador As Integer, ByVal Perdedor As Integer)
    Call SendData(ToAll, Ganador, 0, "||Retos> " & UserList(Ganador).name & " venció a " & UserList(Perdedor).name & " en un reto." & FONTTYPE_TALK)
    Call ResetDuelo(Ganador, Perdedor)
    End Sub
    Public Sub DesconectarDuelo(ByVal Ganador As Integer, ByVal Perdedor As Integer)
    Call SendData(ToAll, Ganador, 0, "||Retos> El reto ha sido cancelado por la desconexión de " & UserList(Perdedor).name & "." & FONTTYPE_TALK)
    Call ResetDuelo(Ganador, Perdedor)
    End Sub
