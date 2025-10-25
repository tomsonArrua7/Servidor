Attribute VB_Name = "modRetos"
Option Explicit

Private Type tJugadores
    ui As Integer
    rounds As Byte
    canMove As Boolean
End Type

Private Type tRetosArena
    InitialPos(1) As WorldPos
    Ocupada As Boolean
    Jugadores(1) As tJugadores
    Countdown As Integer ' segundos restantes para la cuenta atrás
End Type

Private Type tDuelInitPositions
    player1 As WorldPos
    player2 As WorldPos
End Type

Public Const MAX_ARENAS As Byte = 4
Public Const PRECIO_RETO As Long = 0
Private Const MAPA_RETOS As Integer = 35

Public ArenaReto(1 To MAX_ARENAS) As tRetosArena
Public DuelInitPositions(1 To MAX_ARENAS) As tDuelInitPositions

' Función para restaurar todas las estadísticas al máximo
Sub RestoreAllStats(ByVal UserIndex As Integer)
    With UserList(UserIndex).Stats
        .MinHP = .MaxHP
        .MinMAN = .MaxMAN
        .MinSta = .MaxSta
    End With
End Sub

' Función para encontrar un jugador en una posición específica usando MapData
Function FindPlayerAtPosition(pos As WorldPos) As Integer
    If MapData(pos.Map, pos.X, pos.Y).UserIndex > 0 Then
        FindPlayerAtPosition = MapData(pos.Map, pos.X, pos.Y).UserIndex
    Else
        FindPlayerAtPosition = 0
    End If
End Function

' Inicializa las arenas y posiciones de inicio de duelos
Public Sub Reto_ArenasInit()
    Dim i As Long, z As Long
    For i = 1 To MAX_ARENAS
        For z = 0 To 1
            ArenaReto(i).InitialPos(z).Map = MAPA_RETOS
        Next z
    Next i
    
    ' Arena 1
    With ArenaReto(1)
        .InitialPos(0).X = 47
        .InitialPos(0).Y = 42
        .InitialPos(1).X = 62
        .InitialPos(1).Y = 52
    End With
    
    ' Arena 2
    With ArenaReto(2)
        .InitialPos(0).X = 13
        .InitialPos(0).Y = 17
        .InitialPos(1).X = 26
        .InitialPos(1).Y = 27
    End With
    
    ' Arena 3
    With ArenaReto(3)
        .InitialPos(0).X = 73
        .InitialPos(0).Y = 16
        .InitialPos(1).X = 85
        .InitialPos(1).Y = 26
    End With
    
    ' Arena 4
    With ArenaReto(4)
        .InitialPos(0).X = 18
        .InitialPos(0).Y = 79
        .InitialPos(1).X = 32
        .InitialPos(1).Y = 89
    End With
    
    ' Posiciones de inicio de duelos en mapa 1 (ajusta según tus necesidades)
    With DuelInitPositions(1)
        .player1.Map = 1
        .player1.X = 52
        .player1.Y = 33
        .player2.Map = 1
        .player2.X = 53
        .player2.Y = 33
    End With
    With DuelInitPositions(2)
        .player1.Map = 1
        .player1.X = 20
        .player1.Y = 20
        .player2.Map = 1
        .player2.X = 25
        .player2.Y = 25
    End With
    With DuelInitPositions(3)
        .player1.Map = 1
        .player1.X = 30
        .player1.Y = 30
        .player2.Map = 1
        .player2.X = 35
        .player2.Y = 35
    End With
    With DuelInitPositions(4)
        .player1.Map = 1
        .player1.X = 40
        .player1.Y = 40
        .player2.Map = 1
        .player2.X = 45
        .player2.Y = 45
    End With
End Sub

Public Function Reto_ArenaLibre() As Byte
    Dim i As Long
    For i = 1 To MAX_ARENAS
        If ArenaReto(i).Ocupada = False Then Reto_ArenaLibre = i: Exit Function
    Next i
End Function

Public Sub Reto_Inicia(ByVal arena As Byte, ByVal player1 As Integer, ByVal player2 As Integer)
    Dim i As Long
    
    With ArenaReto(arena)
        .Jugadores(0).ui = player1
        .Jugadores(1).ui = player2
        UserList(player1).Oponente = player2
        UserList(player2).Oponente = player1
        
        .Ocupada = True
        
        For i = 0 To 1
            .Jugadores(i).rounds = 0
            .Jugadores(i).canMove = True ' Permitir movimiento al inicio
            UserList(.Jugadores(i).ui).enReto = True
            UserList(.Jugadores(i).ui).Arena_Reto = arena
            UserList(.Jugadores(i).ui).Stats.GLD = UserList(.Jugadores(i).ui).Stats.GLD - PRECIO_RETO
            Call SendUserORO(.Jugadores(i).ui)
            Call WarpUserChar(.Jugadores(i).ui, .InitialPos(i).Map, .InitialPos(i).X, .InitialPos(i).Y, True)
        Next i
    End With
End Sub

Public Sub Reto_Muere(ByVal UserIndex As Integer)
    Dim currInd As Byte, otherInd As Byte, i As Long
    
    If UserList(UserIndex).Arena_Reto = 0 Or UserList(UserIndex).Arena_Reto > MAX_ARENAS Then Exit Sub
    
    With ArenaReto(UserList(UserIndex).Arena_Reto)
        If .Jugadores(0).ui = UserIndex Then currInd = 0: otherInd = 1
        If .Jugadores(1).ui = UserIndex Then currInd = 1: otherInd = 0
        
        .Jugadores(currInd).rounds = .Jugadores(currInd).rounds + 1
        
        If .Jugadores(currInd).rounds < 2 Then ' Continuar si las pérdidas son menores a 2
            Call RevivirUsuarioNPC(UserIndex)
            For i = 0 To 1
                Call RestoreAllStats(.Jugadores(i).ui)
                Call WarpUserChar(.Jugadores(i).ui, .InitialPos(i).Map, .InitialPos(i).X, .InitialPos(i).Y, True)
                .Jugadores(i).canMove = False ' Bloquear movimiento hasta que termine la cuenta atrás
            Next i
            .Countdown = 3 ' Iniciar cuenta atrás
            Call SendData(ToIndex, .Jugadores(0).ui, 0, "||Retos> Siguiente ronda!" & FONTTYPE_INFO)
            Call SendData(ToIndex, .Jugadores(1).ui, 0, "||Retos> Siguiente ronda!" & FONTTYPE_INFO)
        Else
            Call Reto_Termina(.Jugadores(otherInd).ui, .Jugadores(currInd).ui, UserList(UserIndex).Arena_Reto)
        End If
    End With
End Sub

Public Sub Reto_Termina(ByVal Ganador As Integer, ByVal Perdedor As Integer, ByVal arena As Byte)
    UserList(Ganador).flags.Reto = UserList(Ganador).flags.Reto + 1
    UserList(Ganador).Stats.GLD = UserList(Ganador).Stats.GLD + (PRECIO_RETO * 1.5)
    
    Call SendUserORO(Ganador)
    
    Call SendData(ToAll, 0, 0, "||Retos> " & UserList(Ganador).Name & " ha ganado el reto!" & FONTTYPE_RETOS)
    
    UserList(Ganador).enReto = False
    UserList(Perdedor).enReto = False
    
    ArenaReto(arena).Ocupada = False
    ArenaReto(arena).Countdown = 0 ' Resetear cuenta atrás
    
    Call WarpUserChar(Ganador, ULLATHORPE.Map, ULLATHORPE.X, ULLATHORPE.Y, True)
    Call WarpUserChar(Perdedor, ULLATHORPE.Map, ULLATHORPE.X, ULLATHORPE.Y, True)
End Sub

Public Sub CheckDuelPosition(ByVal UserIndex As Integer)
    Dim pos As WorldPos
    Dim i As Integer
    
    pos = UserList(UserIndex).pos
    
    For i = 1 To MAX_ARENAS
        If pos.Map = DuelInitPositions(i).player1.Map And _
           pos.X = DuelInitPositions(i).player1.X And _
           pos.Y = DuelInitPositions(i).player1.Y Then
            Dim player2Index As Integer
            player2Index = FindPlayerAtPosition(DuelInitPositions(i).player2)
            If player2Index > 0 And player2Index <> UserIndex Then
                If Not ArenaReto(i).Ocupada Then
                    Call Reto_Inicia(i, UserIndex, player2Index)
                End If
            End If
        ElseIf pos.Map = DuelInitPositions(i).player2.Map And _
               pos.X = DuelInitPositions(i).player2.X And _
               pos.Y = DuelInitPositions(i).player2.Y Then
            Dim player1Index As Integer
            player1Index = FindPlayerAtPosition(DuelInitPositions(i).player1)
            If player1Index > 0 And player1Index <> UserIndex Then
                If Not ArenaReto(i).Ocupada Then
                    Call Reto_Inicia(i, player1Index, UserIndex)
                End If
            End If
        End If
    Next i
End Sub










