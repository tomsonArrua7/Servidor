Attribute VB_Name = "TCP"
'================================
'>>>>> WWW.FADICTOS.COM.AR <<<<<<
'================================
'Matute - matius_xd07@hotmail.com

Option Explicit
Public usercorreo As String

Public Const SOCKET_BUFFER_SIZE = 3072
Public Enpausa As Boolean

Public Const COMMAND_BUFFER_SIZE = 1000
Public entorneo As Byte

Public Const NingunArma = 2
Dim Response As String
Dim Start As Single, Tmr As Single


Public Const ToIndex = 0
Public Const ToAll = 1
Public Const ToMap = 2
Public Const ToPCArea = 3
Public Const ToNone = 4
Public Const ToAllButIndex = 5
Public Const ToMapButIndex = 6
Public Const ToGM = 7
Public Const ToNPCArea = 8
Public Const ToGuildMembers = 9
Public Const ToAdmins = 10
Public Const ToPCAreaButIndex = 11
Public Const ToMuertos = 12
Public Const ToPCAreaVivos = 13
Public Const ToNPCAreaG = 14
Public Const ToPCAreaButIndexG = 15
Public Const ToGMArea = 16
Public Const ToPCAreaG = 17
Public Const ToAlianza = 18
Public Const ToCaos = 19
Public Const ToParty = 20
Public Const ToMoreAdmins = 21


#If UsarQueSocket = 0 Then
Public Const INVALID_HANDLE = -1
Public Const CONTROL_ERRIGNORE = 0
Public Const CONTROL_ERRDISPLAY = 1



Public Const SOCKET_OPEN = 1
Public Const SOCKET_CONNECT = 2
Public Const SOCKET_LISTEN = 3
Public Const SOCKET_ACCEPT = 4
Public Const SOCKET_CANCEL = 5
Public Const SOCKET_FLUSH = 6
Public Const SOCKET_CLOSE = 7
Public Const SOCKET_DISCONNECT = 7
Public Const SOCKET_ABORT = 8


Public Const SOCKET_NONE = 0
Public Const SOCKET_IDLE = 1
Public Const SOCKET_LISTENING = 2
Public Const SOCKET_CONNECTING = 3
Public Const SOCKET_ACCEPTING = 4
Public Const SOCKET_RECEIVING = 5
Public Const SOCKET_SENDING = 6
Public Const SOCKET_CLOSING = 7


Public Const AF_UNSPEC = 0
Public Const AF_UNIX = 1
Public Const AF_INET = 2


Public Const SOCK_STREAM = 1
Public Const SOCK_DGRAM = 2
Public Const SOCK_RAW = 3
Public Const SOCK_RDM = 4
Public Const SOCK_SEQPACKET = 5


Public Const IPPROTO_IP = 0
Public Const IPPROTO_ICMP = 1
Public Const IPPROTO_GGP = 2
Public Const IPPROTO_TCP = 6
Public Const IPPROTO_PUP = 12
Public Const IPPROTO_UDP = 17
Public Const IPPROTO_IDP = 22
Public Const IPPROTO_ND = 77
Public Const IPPROTO_RAW = 255
Public Const IPPROTO_MAX = 256



Public Const INADDR_ANY = "0.0.0.0"
Public Const INADDR_LOOPBACK = "127.0.0.1"
Public Const INADDR_NONE = "255.055.255.255"


Public Const SOCKET_READ = 0
Public Const SOCKET_WRITE = 1
Public Const SOCKET_READWRITE = 2


Public Const SOCKET_ERRIGNORE = 0
Public Const SOCKET_ERRDISPLAY = 1


Public Const WSABASEERR = 24000
Public Const WSAEINTR = 24004
Public Const WSAEBADF = 24009
Public Const WSAEACCES = 24013
Public Const WSAEFAULT = 24014
Public Const WSAEINVAL = 24022
Public Const WSAEMFILE = 24024
Public Const WSAEWOULDBLOCK = 24035
Public Const WSAEINPROGRESS = 24036
Public Const WSAEALREADY = 24037
Public Const WSAENOTSOCK = 24038
Public Const WSAEDESTADDRREQ = 24039
Public Const WSAEMSGSIZE = 24040
Public Const WSAEPROTOTYPE = 24041
Public Const WSAENOPROTOOPT = 24042
Public Const WSAEPROTONOSUPPORT = 24043
Public Const WSAESOCKTNOSUPPORT = 24044
Public Const WSAEOPNOTSUPP = 24045
Public Const WSAEPFNOSUPPORT = 24046
Public Const WSAEAFNOSUPPORT = 24047
Public Const WSAEADDRINUSE = 24048
Public Const WSAEADDRNOTAVAIL = 24049
Public Const WSAENETDOWN = 24050
Public Const WSAENETUNREACH = 24051
Public Const WSAENETRESET = 24052
Public Const WSAECONNABORTED = 24053
Public Const WSAECONNRESET = 24054
Public Const WSAENOBUFS = 24055
Public Const WSAEISCONN = 24056
Public Const WSAENOTCONN = 24057
Public Const WSAESHUTDOWN = 24058
Public Const WSAETOOMANYREFS = 24059
Public Const WSAETIMEDOUT = 24060
Public Const WSAECONNREFUSED = 24061
Public Const WSAELOOP = 24062
Public Const WSAENAMETOOLONG = 24063
Public Const WSAEHOSTDOWN = 24064
Public Const WSAEHOSTUNREACH = 24065
Public Const WSAENOTEMPTY = 24066
Public Const WSAEPROCLIM = 24067
Public Const WSAEUSERS = 24068
Public Const WSAEDQUOT = 24069
Public Const WSAESTALE = 24070
Public Const WSAEREMOTE = 24071
Public Const WSASYSNOTREADY = 24091
Public Const WSAVERNOTSUPPORTED = 24092
Public Const WSANOTINITIALISED = 24093
Public Const WSAHOST_NOT_FOUND = 25001
Public Const WSATRY_AGAIN = 25002
Public Const WSANO_RECOVERY = 25003
Public Const WSANO_DATA = 25004
Public Const WSANO_ADDRESS = 2500
#End If

Public Data(1 To 3, 1 To 2, 1 To 2, 1 To 2) As Double
Public Onlines(1 To 3) As Long

Public Const Minuto = 1
Public Const Hora = 2
Public Const Dia = 3

Public Const Actual = 1
Public Const Last = 2

Public Const Enviada = 1
Public Const Recibida = 2

Public Const Mensages = 1
Public Const Letras = 2

Sub DarCuerpoYCabeza(UserBody As Integer, UserHead As Integer, Raza As Byte, Gen As Byte)

Select Case Gen
   Case HOMBRE
        Select Case Raza
        
                Case HUMANO
                    UserHead = CInt(RandomNumber(1, 24))
                    If UserHead > 24 Then UserHead = 24
                    UserBody = 1
                Case ELFO
                    UserHead = CInt(RandomNumber(1, 7)) + 100
                    If UserHead > 107 Then UserHead = 107
                    UserBody = 2
                Case ELFO_OSCURO
                    UserHead = CInt(RandomNumber(1, 4)) + 200
                    If UserHead > 204 Then UserHead = 204
                    UserBody = 3
                Case ENANO
                    UserHead = RandomNumber(1, 4) + 300
                    If UserHead > 304 Then UserHead = 304
                    UserBody = 52
                Case GNOMO
                    UserHead = RandomNumber(1, 3) + 400
                    If UserHead > 403 Then UserHead = 403
                    UserBody = 52
                Case Else
                    UserHead = 1
                    UserBody = 1
            
        End Select
   Case MUJER
        Select Case Raza
                Case HUMANO
                    UserHead = CInt(RandomNumber(1, 4)) + 69
                    If UserHead > 73 Then UserHead = 73
                    UserBody = 1
                Case ELFO
                    UserHead = CInt(RandomNumber(1, 5)) + 169
                    If UserHead > 174 Then UserHead = 174
                    UserBody = 2
                Case ELFO_OSCURO
                    UserHead = CInt(RandomNumber(1, 5)) + 269
                    If UserHead > 274 Then UserHead = 274
                    UserBody = 3
                Case GNOMO
                    UserHead = RandomNumber(1, 4) + 469
                    If UserHead > 473 Then UserHead = 473
                    UserBody = 52
                Case ENANO
                    UserHead = RandomNumber(1, 3) + 369
                    If UserHead > 372 Then UserHead = 372
                    UserBody = 52
                Case Else
                    UserHead = 70
                    UserBody = 1
        End Select
End Select

   
End Sub
Function AsciiValidos(ByVal cad As String) As Boolean
Dim car As Byte
Dim i As Integer

cad = LCase$(cad)

For i = 1 To Len(cad)
    car = Asc(Mid$(cad, i, 1))
    If (car < 97 Or car > 122) And (car <> 255) And (car <> 32) Then
        AsciiValidos = False
        Exit Function
    End If
Next

AsciiValidos = True

End Function

Function Numeric(ByVal cad As String) As Boolean
Dim car As Byte
Dim i As Integer

cad = LCase$(cad)

For i = 1 To Len(cad)
    car = Asc(Mid$(cad, i, 1))
    
    If (car < 48 Or car > 57) Then
        Numeric = False
        Exit Function
    End If
    
Next

Numeric = True

End Function
Function NombrePermitido(ByVal Nombre As String) As Boolean
Dim i As Integer

For i = 1 To UBound(ForbidenNames)
    If InStr(Nombre, ForbidenNames(i)) Then
        NombrePermitido = False
        Exit Function
    End If
Next

NombrePermitido = True

End Function

Function ValidateAtrib(UserIndex As Integer) As Boolean
Dim LoopC As Integer

For LoopC = 1 To NUMATRIBUTOS
    If UserList(UserIndex).Stats.UserAtributosBackUP(LoopC) > 23 Or UserList(UserIndex).Stats.UserAtributosBackUP(LoopC) < 1 Then Exit Function
Next

ValidateAtrib = True

End Function

Function ValidateAtrib2(UserIndex As Integer) As Boolean
Dim LoopC As Integer

For LoopC = 1 To NUMATRIBUTOS
    If UserList(UserIndex).Stats.UserAtributosBackUP(LoopC) > 18 Or UserList(UserIndex).Stats.UserAtributosBackUP(LoopC) < 1 Then
    ValidateAtrib2 = False
    Exit Function
    End If
Next

ValidateAtrib2 = True

End Function
Function ValidateSkills(UserIndex As Integer) As Boolean
Dim LoopC As Integer

For LoopC = 1 To NUMSKILLS
    If UserList(UserIndex).Stats.UserSkills(LoopC) < 0 Then Exit Function
    If UserList(UserIndex).Stats.UserSkills(LoopC) > 100 Then UserList(UserIndex).Stats.UserSkills(LoopC) = 100
Next

ValidateSkills = True

End Function
Sub ConnectNewUser(UserIndex As Integer, Name As String, Password As String, _
Body As Integer, Head As Integer, UserRaza As Byte, UserSexo As Byte, _
UA1 As String, UA2 As String, UA3 As String, UA4 As String, UA5 As String, _
US1 As String, US2 As String, US3 As String, US4 As String, US5 As String, _
US6 As String, US7 As String, US8 As String, US9 As String, US10 As String, _
US11 As String, US12 As String, US13 As String, US14 As String, US15 As String, _
US16 As String, US17 As String, US18 As String, US19 As String, US20 As String, _
US21 As String, US22 As String, UserEmail As String, Hogar As Byte)

Dim i As Integer

If Restringido Then
    Call SendData(ToIndex, UserIndex, 0, "ERREl servidor está restringido solo para GameMasters temporalmente.")
    Exit Sub
End If

If Not NombrePermitido(Name) Then
    Call SendData(ToIndex, UserIndex, 0, "ERRLos nombres de los personajes deben pertencer a la fantasia, el nombre indicado es invalido.")
    Call SendData(ToIndex, UserIndex, 0, "V8V" & 2)
    Exit Sub
End If

If Not AsciiValidos(Name) Then
    Call SendData(ToIndex, UserIndex, 0, "ERRNombre invalido.")
    Call SendData(ToIndex, UserIndex, 0, "V8V" & 2)
    Exit Sub
End If

Dim LoopC As Integer
Dim totalskpts As Long
  

'¿Existe el personaje?
If FileExist(CharPath & UCase$(Name) & ".chr", vbNormal) = True Then
    Call SendData(ToIndex, UserIndex, 0, "ERRYa existe el personaje.")
    Exit Sub
End If

UserList(UserIndex).flags.Muerto = 0
UserList(UserIndex).flags.Escondido = 0

UserList(UserIndex).Name = Name
UserList(UserIndex).Clase = CIUDADANO
UserList(UserIndex).Raza = UserRaza
UserList(UserIndex).Genero = UserSexo
UserList(UserIndex).Email = UserEmail
UserList(UserIndex).Hogar = Hogar

Select Case UserList(UserIndex).Raza
    Case HUMANO
        UserList(UserIndex).Stats.UserAtributosBackUP(fuerza) = UserList(UserIndex).Stats.UserAtributosBackUP(fuerza) + 1
        UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) = UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) + 1
        UserList(UserIndex).Stats.UserAtributosBackUP(Constitucion) = UserList(UserIndex).Stats.UserAtributosBackUP(Constitucion) + 2
    Case ELFO
        UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) = UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) + 3
        UserList(UserIndex).Stats.UserAtributosBackUP(Constitucion) = UserList(UserIndex).Stats.UserAtributosBackUP(Constitucion) + 1
        UserList(UserIndex).Stats.UserAtributosBackUP(Inteligencia) = UserList(UserIndex).Stats.UserAtributosBackUP(Inteligencia) + 1
        UserList(UserIndex).Stats.UserAtributosBackUP(Carisma) = UserList(UserIndex).Stats.UserAtributosBackUP(Carisma) + 2
    Case ELFO_OSCURO
        UserList(UserIndex).Stats.UserAtributosBackUP(fuerza) = UserList(UserIndex).Stats.UserAtributosBackUP(fuerza) + 1
        UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) = UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) + 1
        UserList(UserIndex).Stats.UserAtributosBackUP(Carisma) = UserList(UserIndex).Stats.UserAtributosBackUP(Carisma) - 3
        UserList(UserIndex).Stats.UserAtributosBackUP(Inteligencia) = UserList(UserIndex).Stats.UserAtributosBackUP(Inteligencia) + 2
    Case ENANO
        UserList(UserIndex).Stats.UserAtributosBackUP(fuerza) = UserList(UserIndex).Stats.UserAtributosBackUP(fuerza) + 3
        UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) = UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) - 1
        UserList(UserIndex).Stats.UserAtributosBackUP(Constitucion) = UserList(UserIndex).Stats.UserAtributosBackUP(Constitucion) + 3
        UserList(UserIndex).Stats.UserAtributosBackUP(Inteligencia) = UserList(UserIndex).Stats.UserAtributosBackUP(Inteligencia) - 6
        UserList(UserIndex).Stats.UserAtributosBackUP(Carisma) = UserList(UserIndex).Stats.UserAtributosBackUP(Carisma) - 3
    Case GNOMO
        UserList(UserIndex).Stats.UserAtributosBackUP(fuerza) = UserList(UserIndex).Stats.UserAtributosBackUP(fuerza) - 5
        UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) = UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) + 4
        UserList(UserIndex).Stats.UserAtributosBackUP(Inteligencia) = UserList(UserIndex).Stats.UserAtributosBackUP(Inteligencia) + 3
        UserList(UserIndex).Stats.UserAtributosBackUP(Carisma) = UserList(UserIndex).Stats.UserAtributosBackUP(Carisma) + 1
End Select

If Not ValidateAtrib(UserIndex) Then
    Call SendData(ToIndex, UserIndex, 0, "ERRAtributos invalidos.")
    Call SendData(ToIndex, UserIndex, 0, "V8V" & 2)
    Exit Sub
End If

UserList(UserIndex).Stats.UserSkills(1) = val(US1)
UserList(UserIndex).Stats.UserSkills(2) = val(US2)
UserList(UserIndex).Stats.UserSkills(3) = val(US3)
UserList(UserIndex).Stats.UserSkills(4) = val(US4)
UserList(UserIndex).Stats.UserSkills(5) = val(US5)
UserList(UserIndex).Stats.UserSkills(6) = val(US6)
UserList(UserIndex).Stats.UserSkills(7) = val(US7)
UserList(UserIndex).Stats.UserSkills(8) = val(US8)
UserList(UserIndex).Stats.UserSkills(9) = val(US9)
UserList(UserIndex).Stats.UserSkills(10) = val(US10)
UserList(UserIndex).Stats.UserSkills(11) = val(US11)
UserList(UserIndex).Stats.UserSkills(12) = val(US12)
UserList(UserIndex).Stats.UserSkills(13) = val(US13)
UserList(UserIndex).Stats.UserSkills(14) = val(US14)
UserList(UserIndex).Stats.UserSkills(15) = val(US15)
UserList(UserIndex).Stats.UserSkills(16) = val(US16)
UserList(UserIndex).Stats.UserSkills(17) = val(US17)
UserList(UserIndex).Stats.UserSkills(18) = val(US18)
UserList(UserIndex).Stats.UserSkills(19) = val(US19)
UserList(UserIndex).Stats.UserSkills(20) = val(US20)
UserList(UserIndex).Stats.UserSkills(21) = val(US21)
UserList(UserIndex).Stats.UserSkills(22) = val(US22)

totalskpts = 0


For LoopC = 1 To NUMSKILLS
    totalskpts = totalskpts + Abs(UserList(UserIndex).Stats.UserSkills(LoopC))
Next

miuseremail = UserEmail
If totalskpts > 10 Then
    Call LogHackAttemp(UserList(UserIndex).Name & " intento hackear los skills.")
  
    Call CloseSocket(UserIndex)
    Exit Sub
End If


UserList(UserIndex).Password = Password

UserList(UserIndex).Char.Heading = SOUTH

Call DarCuerpoYCabeza(UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Raza, UserList(UserIndex).Genero)
UserList(UserIndex).OrigChar = UserList(UserIndex).Char
   
UserList(UserIndex).Char.WeaponAnim = NingunArma
UserList(UserIndex).Char.ShieldAnim = NingunEscudo
UserList(UserIndex).Char.CascoAnim = NingunCasco

UserList(UserIndex).Stats.MET = 1
Dim MiInt
MiInt = RandomNumber(1, UserList(UserIndex).Stats.UserAtributosBackUP(Constitucion) \ 3)

UserList(UserIndex).Stats.MaxHP = 15 + MiInt
UserList(UserIndex).Stats.MinHP = 15 + MiInt

UserList(UserIndex).Stats.FIT = 1


MiInt = RandomNumber(1, UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) \ 6)
If MiInt = 1 Then MiInt = 2

UserList(UserIndex).Stats.MaxSta = 20 * MiInt
UserList(UserIndex).Stats.MinSta = 20 * MiInt

UserList(UserIndex).Stats.MaxAGU = 100
UserList(UserIndex).Stats.MinAGU = 100

UserList(UserIndex).Stats.MaxHam = 100
UserList(UserIndex).Stats.MinHam = 100




    UserList(UserIndex).Stats.MaxMAN = 0
    UserList(UserIndex).Stats.MinMAN = 0


UserList(UserIndex).Stats.MaxHit = 2
UserList(UserIndex).Stats.MinHit = 1

UserList(UserIndex).Stats.GLD = 0




UserList(UserIndex).Stats.Exp = 0
UserList(UserIndex).Stats.ELU = ELUs(1)
UserList(UserIndex).Stats.ELV = 1



UserList(UserIndex).Invent.NroItems = 4

UserList(UserIndex).Invent.Object(1).OBJIndex = ManzanaNewbie
UserList(UserIndex).Invent.Object(1).Amount = 40

UserList(UserIndex).Invent.Object(2).OBJIndex = AguaNewbie
UserList(UserIndex).Invent.Object(2).Amount = 40

UserList(UserIndex).Invent.Object(3).OBJIndex = DagaNewbie
UserList(UserIndex).Invent.Object(3).Amount = 1
UserList(UserIndex).Invent.Object(3).Equipped = 1

Select Case UserList(UserIndex).Raza
    Case HUMANO
        UserList(UserIndex).Invent.Object(4).OBJIndex = RopaNewbieHumano
    Case ELFO
        UserList(UserIndex).Invent.Object(4).OBJIndex = RopaNewbieElfo
    Case ELFO_OSCURO
        UserList(UserIndex).Invent.Object(4).OBJIndex = RopaNewbieElfoOscuro
    Case Else
        UserList(UserIndex).Invent.Object(4).OBJIndex = RopaNewbieEnano
End Select

UserList(UserIndex).Invent.Object(4).Amount = 1
UserList(UserIndex).Invent.Object(4).Equipped = 1

UserList(UserIndex).Invent.Object(5).OBJIndex = PocionRojaNewbie
UserList(UserIndex).Invent.Object(5).Amount = 500

UserList(UserIndex).Invent.Object(6).OBJIndex = PocionVerdeNewbie
UserList(UserIndex).Invent.Object(6).Amount = 30

UserList(UserIndex).Invent.ArmourEqpSlot = 4
UserList(UserIndex).Invent.ArmourEqpObjIndex = UserList(UserIndex).Invent.Object(4).OBJIndex

UserList(UserIndex).Invent.WeaponEqpObjIndex = UserList(UserIndex).Invent.Object(3).OBJIndex
UserList(UserIndex).Invent.WeaponEqpSlot = 3

Call SaveUser(UserIndex, CharPath & UCase$(Name) & ".chr")
Call ConnectUser(UserIndex, Name, Password)

End Sub
Sub CloseSocket(ByVal UserIndex As Integer, Optional ByVal cerrarlo As Boolean = True)
On Error GoTo errhandler
Dim LoopC As Integer

Call aDos.RestarConexion(UserList(UserIndex).ip)

Call VerificarRetos(UserIndex)
If UserList(UserIndex).flags.UserLogged Then
    If NumUsers > 0 Then NumUsers = NumUsers - 1
    If UserList(UserIndex).flags.Privilegios = 0 Then NumNoGMs = NumNoGMs - 1
    Call EstadisticasWeb.Informar(CANTIDAD_ONLINE, NumUsers)
    Call CloseUser(UserIndex)
End If

If UserList(UserIndex).flags.automatico = True Then
Call Rondas_UsuarioDesconecta(UserIndex)
End If

If UserList(UserIndex).ConnID <> -1 Then Call ApiCloseSocket(UserList(UserIndex).ConnID)

UserList(UserIndex) = UserOffline

Exit Sub

errhandler:
    UserList(UserIndex) = UserOffline
    Call LogError("Error en CloseSocket " & Err.Description)

End Sub

Sub SendData(sndRoute As Byte, sndIndex As Integer, sndMap As Integer, sndData As String)
Dim LoopC As Integer
Dim aux$
Dim dec$
Dim nfile As Integer
Dim Ret As Long

sndData = sndData & ENDC

Select Case sndRoute

    Case ToIndex
        If UserList(sndIndex).ConnID > -1 Then
             Call WsApiEnviar(sndIndex, sndData)
             Exit Sub
        End If
        Exit Sub

    Case ToMap
        
        For LoopC = 1 To MapInfo(sndMap).NumUsers
            Call WsApiEnviar(MapInfo(sndMap).UserIndex(LoopC), sndData)
        Next
        Exit Sub

    Case ToPCArea
        
        
        For LoopC = 1 To MapInfo(UserList(sndIndex).POS.Map).NumUsers
            If EnPantalla(UserList(sndIndex).POS, UserList(MapInfo(UserList(sndIndex).POS.Map).UserIndex(LoopC)).POS, 1) Then Call WsApiEnviar(MapInfo(UserList(sndIndex).POS.Map).UserIndex(LoopC), sndData)
        Next
        Exit Sub
        
    Case ToNone
        Exit Sub

    Case ToAdmins
        For LoopC = 1 To LastUser
            If UserList(LoopC).ConnID > -1 And UserList(LoopC).flags.Privilegios Then Call WsApiEnviar(LoopC, sndData)
        Next
        Exit Sub
        
    Case ToMoreAdmins
        For LoopC = 1 To LastUser
            If UserList(LoopC).ConnID > -1 And UserList(LoopC).flags.Privilegios >= UserList(sndIndex).flags.Privilegios Then Call WsApiEnviar(LoopC, sndData)
        Next
        Exit Sub
        
    Case ToParty
        Dim MiembroIndex As Integer
        If UserList(sndIndex).PartyIndex = 0 Then Exit Sub
        For LoopC = 1 To MAXPARTYUSERS
            MiembroIndex = Party(UserList(sndIndex).PartyIndex).MiembrosIndex(LoopC)
            If MiembroIndex > 0 Then
                If UserList(MiembroIndex).ConnID > -1 And UserList(MiembroIndex).flags.UserLogged And UserList(MiembroIndex).flags.Party > 0 Then Call WsApiEnviar(MiembroIndex, sndData)
            End If
        Next
        
        Exit Sub
        
    Case ToAll
        For LoopC = 1 To LastUser
            If UserList(LoopC).ConnID > -1 And UserList(LoopC).flags.UserLogged Then Call WsApiEnviar(LoopC, sndData)
        Next
        Exit Sub
    
    Case ToAllButIndex
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID > -1) And (LoopC <> sndIndex) And UserList(LoopC).flags.UserLogged Then Call WsApiEnviar(LoopC, sndData)
        Next
        Exit Sub
      
    Case ToMapButIndex
        For LoopC = 1 To MapInfo(UserList(sndIndex).POS.Map).NumUsers
            If MapInfo(UserList(sndIndex).POS.Map).UserIndex(LoopC) <> sndIndex Then Call WsApiEnviar(MapInfo(UserList(sndIndex).POS.Map).UserIndex(LoopC), sndData)
        Next
        Exit Sub
            
    Case ToGuildMembers
        If Len(UserList(sndIndex).GuildInfo.GuildName) = 0 Then Exit Sub
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID > -1) And UserList(sndIndex).GuildInfo.GuildName = UserList(LoopC).GuildInfo.GuildName Then Call WsApiEnviar(LoopC, sndData)
        Next
        Exit Sub
        
    Case ToGMArea
        For LoopC = 1 To MapInfo(UserList(sndIndex).POS.Map).NumUsers
            If EnPantalla(UserList(sndIndex).POS, UserList(MapInfo(UserList(sndIndex).POS.Map).UserIndex(LoopC)).POS, 1) And UserList(MapInfo(UserList(sndIndex).POS.Map).UserIndex(LoopC)).flags.Privilegios Then Call WsApiEnviar(MapInfo(UserList(sndIndex).POS.Map).UserIndex(LoopC), sndData)
        Next
        Exit Sub

    Case ToPCAreaVivos
        For LoopC = 1 To MapInfo(UserList(sndIndex).POS.Map).NumUsers
            If EnPantalla(UserList(sndIndex).POS, UserList(MapInfo(UserList(sndIndex).POS.Map).UserIndex(LoopC)).POS, 1) Then
                If Not UserList(MapInfo(UserList(sndIndex).POS.Map).UserIndex(LoopC)).flags.Muerto Or UserList(MapInfo(UserList(sndIndex).POS.Map).UserIndex(LoopC)).Clase = CLERIGO Then Call WsApiEnviar(MapInfo(UserList(sndIndex).POS.Map).UserIndex(LoopC), sndData)
            End If
        Next
        Exit Sub
        
    Case ToMuertos
        For LoopC = 1 To MapInfo(UserList(sndIndex).POS.Map).NumUsers
            If EnPantalla(UserList(sndIndex).POS, UserList(MapInfo(UserList(sndIndex).POS.Map).UserIndex(LoopC)).POS, 1) Then
                If UserList(MapInfo(UserList(sndIndex).POS.Map).UserIndex(LoopC)).Clase = CLERIGO Or UserList(MapInfo(UserList(sndIndex).POS.Map).UserIndex(LoopC)).flags.Muerto Or UserList(MapInfo(UserList(sndIndex).POS.Map).UserIndex(LoopC)).flags.Privilegios Then Call WsApiEnviar(MapInfo(UserList(sndIndex).POS.Map).UserIndex(LoopC), sndData)
            End If
        Next
        Exit Sub

    Case ToPCAreaButIndex
        For LoopC = 1 To MapInfo(UserList(sndIndex).POS.Map).NumUsers
            If EnPantalla(UserList(sndIndex).POS, UserList(MapInfo(UserList(sndIndex).POS.Map).UserIndex(LoopC)).POS, 1) And MapInfo(UserList(sndIndex).POS.Map).UserIndex(LoopC) <> sndIndex Then Call WsApiEnviar(MapInfo(UserList(sndIndex).POS.Map).UserIndex(LoopC), sndData)
        Next
        Exit Sub
        
    Case ToPCAreaButIndexG
        For LoopC = 1 To MapInfo(UserList(sndIndex).POS.Map).NumUsers
            If EnPantalla(UserList(sndIndex).POS, UserList(MapInfo(UserList(sndIndex).POS.Map).UserIndex(LoopC)).POS, 3) And MapInfo(UserList(sndIndex).POS.Map).UserIndex(LoopC) <> sndIndex Then Call WsApiEnviar(MapInfo(UserList(sndIndex).POS.Map).UserIndex(LoopC), sndData)
        Next
        Exit Sub
        
    Case ToNPCArea
        For LoopC = 1 To MapInfo(Npclist(sndIndex).POS.Map).NumUsers
            If EnPantalla(Npclist(sndIndex).POS, UserList(MapInfo(Npclist(sndIndex).POS.Map).UserIndex(LoopC)).POS, 1) Then Call WsApiEnviar(MapInfo(Npclist(sndIndex).POS.Map).UserIndex(LoopC), sndData)
        Next
        Exit Sub

    Case ToNPCAreaG
        For LoopC = 1 To MapInfo(Npclist(sndIndex).POS.Map).NumUsers
            If EnPantalla(Npclist(sndIndex).POS, UserList(MapInfo(Npclist(sndIndex).POS.Map).UserIndex(LoopC)).POS, 3) Then Call WsApiEnviar(MapInfo(Npclist(sndIndex).POS.Map).UserIndex(LoopC), sndData)
        Next
        Exit Sub
        
    Case ToPCAreaG
        For LoopC = 1 To MapInfo(UserList(sndIndex).POS.Map).NumUsers
            If EnPantalla(UserList(sndIndex).POS, UserList(MapInfo(UserList(sndIndex).POS.Map).UserIndex(LoopC)).POS, 3) Then Call WsApiEnviar(MapInfo(UserList(sndIndex).POS.Map).UserIndex(LoopC), sndData)
        Next
        Exit Sub
        
    Case ToAlianza
        For LoopC = 1 To LastUser
            If UserList(LoopC).ConnID > -1 And UserList(LoopC).Faccion.Bando = Real Then Call WsApiEnviar(LoopC, sndData)
        Next
        Exit Sub
        
    Case ToCaos
        For LoopC = 1 To LastUser
            If UserList(LoopC).ConnID > -1 And UserList(LoopC).Faccion.Bando = Caos Then Call WsApiEnviar(LoopC, sndData)
        Next
        Exit Sub

End Select

Exit Sub
Error:
    Call LogError("Error en SendData: " & sndData & "-" & Err.Description & "-Ruta: " & sndRoute & "-Index:" & sndIndex & "-Mapa" & sndMap)
    
End Sub
Function HayPCarea(POS As WorldPos) As Boolean
Dim i As Integer

For i = 1 To MapInfo(POS.Map).NumUsers
    If EnPantalla(POS, UserList(MapInfo(POS.Map).UserIndex(i)).POS, 1) Then
        HayPCarea = True
        Exit Function
    End If
Next

End Function
Function HayOBJarea(POS As WorldPos, OBJIndex As Integer) As Boolean
Dim X As Integer, Y As Integer

For Y = POS.Y - MinYBorder + 1 To POS.Y + MinYBorder - 1
    For X = POS.X - MinXBorder + 1 To POS.X + MinXBorder - 1
        If MapData(POS.Map, X, Y).OBJInfo.OBJIndex = OBJIndex Then
            HayOBJarea = True
            Exit Function
        End If
    Next
Next

End Function

Sub CorregirSkills(UserIndex As Integer)
Dim k As Integer

For k = 1 To NUMSKILLS
  If UserList(UserIndex).Stats.UserSkills(k) > MAXSKILLPOINTS Then UserList(UserIndex).Stats.UserSkills(k) = MAXSKILLPOINTS
Next

For k = 1 To NUMATRIBUTOS
 If UserList(UserIndex).Stats.UserAtributos(k) > MAXATRIBUTOS Then
    Call SendData(ToIndex, UserIndex, 0, "ERREl personaje tiene atributos invalidos.")
    Exit Sub
 End If
Next
 
End Sub
Function ValidateChr(UserIndex As Integer) As Boolean

ValidateChr = (UserList(UserIndex).Char.Head <> 0 Or UserList(UserIndex).flags.Navegando = 1) And _
UserList(UserIndex).Char.Body <> 0 And ValidateSkills(UserIndex)

End Function
Sub ConnectUser(UserIndex As Integer, Name As String, Password As String)
On Error GoTo Error
Dim Privilegios As Byte
Dim N As Integer
Dim LoopC As Integer
Dim o As Integer

UserList(UserIndex).Counters.Protegido = 4
UserList(UserIndex).flags.Protegido = 2

Dim numeromail As Integer

If NumUsers > MaxUsers2 Then
    If Not (EsDios(Name) Or EsSemiDios(Name)) Then
        Call SendData(ToIndex, UserIndex, 0, "ERREl servidor ha alcanzado el maximo de usuarios soportado, por favor vuelva a intertarlo más tarde.")
        Exit Sub
    End If
End If

If NumUsers >= MaxUsers Then
    Call SendData(ToIndex, UserIndex, 0, "ERRLímite de usuarios alcanzado.")
    Call CloseSocket(UserIndex)
    Exit Sub
End If

If AllowMultiLogins = 0 Then
    If CheckForSameIP(UserIndex, UserList(UserIndex).ip) Then
        Call SendData(ToIndex, UserIndex, 0, "ERRNo es posible usar más de un personaje al mismo tiempo.")
        Call CloseSocket(UserIndex)
        Exit Sub
    End If
End If

If CheckForSameName(UserIndex, Name) Then
    If NameIndex(Name) = UserIndex Then Call CloseSocket(NameIndex(Name))
    Call SendData(ToIndex, UserIndex, 0, "ERRPerdón, un usuario con el mismo nombre se ha logeado.")
    Call CloseSocket(UserIndex)
    Exit Sub
End If

'¿Existe el personaje?
If FileExist(CharPath & UCase$(Name) & ".chr", vbNormal) = False Then
    Call SendData(ToIndex, UserIndex, 0, "ERREl personaje no existe.")
    Call CloseSocket(UserIndex)
    Exit Sub
End If

'¿Es el passwd valido?
If UCase$(Password) <> UCase$(GetVar(CharPath & UCase$(Name) & ".chr", "INIT", "Password")) Then
    Call SendData(ToIndex, UserIndex, 0, "ERRPassword incorrecto.")
    Call CloseSocket(UserIndex)
    Exit Sub
End If

If BANCheck(Name) Then
    For LoopC = 1 To Baneos.Count
        If Baneos(LoopC).Name = UCase$(Name) Then
            Call SendData(ToIndex, UserIndex, 0, "ERRSe te ha prohibido la entrada a FénixAO hasta el día " & Format(Baneos(LoopC).FechaLiberacion, "dddddd") & " a las " & Format(Baneos(LoopC).FechaLiberacion, "hh:mm am/pm"))
            Exit Sub
        End If
    Next
    Call SendData(ToIndex, UserIndex, 0, "ERRSe te ha prohibido la entrada a FénixAO.")
    Exit Sub
End If

Call LoadUser(UserIndex, CharPath & UCase$(Name) & ".chr")
If EsAdministrador(Name) Then
    Privilegios = 4
    Call LogGM(Name, "Se conecto con ip:" & UserList(UserIndex).ip, False)
    ElseIf EsDios(Name) Then
    Privilegios = 3
    Call LogGM(Name, "Se conecto con ip:" & UserList(UserIndex).ip, False)
ElseIf EsSemiDios(Name) Then
    Privilegios = 2
    Call LogGM(Name, "Se conecto con ip:" & UserList(UserIndex).ip, False)
ElseIf EsConsejero(Name) Then
    Privilegios = 1
    Call LogGM(Name, "Se conecto con ip:" & UserList(UserIndex).ip, True)
End If

If Restringido And Privilegios = 0 Then
    If Not PuedeDenunciar(Name) Then
        Call SendData(ToIndex, UserIndex, 0, "ERREl servidor está restringido solo para GameMasters temporalmente.")
        Exit Sub
    End If
End If

Dim Quest As Boolean
Quest = PJQuest(Name)

UserList(UserIndex).Counters.IdleCount = Timer
If UserList(UserIndex).Counters.TiempoPena Then UserList(UserIndex).Counters.Pena = Timer
If UserList(UserIndex).flags.Envenenado Then UserList(UserIndex).Counters.Veneno = Timer
UserList(UserIndex).Counters.AGUACounter = Timer
UserList(UserIndex).Counters.COMCounter = Timer

If Not ValidateChr(UserIndex) Then
    Call SendData(ToIndex, UserIndex, 0, "ERRError en el personaje.")
    Call CloseSocket(UserIndex)
    Exit Sub
End If

For o = 1 To BanIps.Count
    If BanIps.Item(o) = UserList(UserIndex).ip Then
        Call CloseSocket(UserIndex)
        Exit Sub
    End If
Next

If UserList(UserIndex).Invent.EscudoEqpSlot = 0 Then UserList(UserIndex).Char.ShieldAnim = NingunEscudo
If UserList(UserIndex).Invent.CascoEqpSlot = 0 Then UserList(UserIndex).Char.CascoAnim = NingunCasco
If UserList(UserIndex).Invent.WeaponEqpSlot = 0 Then UserList(UserIndex).Char.WeaponAnim = NingunArma

Call UpdateUserInv(True, UserIndex, 0)
Call UpdateUserHechizos(True, UserIndex, 0)

If UserList(UserIndex).flags.Navegando = 1 Then
    If UserList(UserIndex).flags.Muerto = 1 Then
        UserList(UserIndex).Char.Body = iFragataFantasmal
        UserList(UserIndex).Char.Head = 0
        UserList(UserIndex).Char.WeaponAnim = NingunArma
        UserList(UserIndex).Char.ShieldAnim = NingunEscudo
        UserList(UserIndex).Char.CascoAnim = NingunCasco
    Else
        UserList(UserIndex).Char.Body = ObjData(UserList(UserIndex).Invent.BarcoObjIndex).Ropaje
        UserList(UserIndex).Char.Head = 0
        UserList(UserIndex).Char.WeaponAnim = NingunArma
        UserList(UserIndex).Char.ShieldAnim = NingunEscudo
        UserList(UserIndex).Char.CascoAnim = NingunCasco
    End If
End If

UserList(UserIndex).flags.Privilegios = Privilegios
UserList(UserIndex).flags.PuedeDenunciar = PuedeDenunciar(Name)
UserList(UserIndex).flags.Quest = Quest

If UserList(UserIndex).flags.Privilegios > 1 Then
    If UCase$(Name) = "BALEY" Then
        UserList(UserIndex).flags.AdminInvisible = 1
        UserList(UserIndex).flags.Invisible = 1
    Else
        UserList(UserIndex).POS.Map = 86
        UserList(UserIndex).POS.X = 50
        UserList(UserIndex).POS.Y = 50
    End If
End If

If UserList(UserIndex).flags.Paralizado Then Call SendData(ToIndex, UserIndex, 0, "P9")
Call SendData(ToIndex, UserIndex, 0, "RPU")
If UserList(UserIndex).POS.Map = 0 Or UserList(UserIndex).POS.Map > NumMaps Then
    Select Case UserList(UserIndex).Hogar
        Case HOGAR_NIX
           UserList(UserIndex).POS.Map = 1
        UserList(UserIndex).POS.X = 50
        UserList(UserIndex).POS.Y = 50
        Case HOGAR_BANDERBILL
            UserList(UserIndex).POS.Map = 1
        UserList(UserIndex).POS.X = 50
        UserList(UserIndex).POS.Y = 50
        Case HOGAR_LINDOS
            UserList(UserIndex).POS.Map = 1
        UserList(UserIndex).POS.X = 50
        UserList(UserIndex).POS.Y = 50
        Case HOGAR_ARGHAL
            UserList(UserIndex).POS.Map = 1
        UserList(UserIndex).POS.X = 50
        UserList(UserIndex).POS.Y = 50
        Case Else
            UserList(UserIndex).POS.Map = 1
        UserList(UserIndex).POS.X = 50
        UserList(UserIndex).POS.Y = 50
    End Select
    If UserList(UserIndex).POS.Map > NumMaps Then UserList(UserIndex).POS = ULLATHORPE
End If

    If MapData(UserList(UserIndex).POS.Map, UserList(UserIndex).POS.X, UserList(UserIndex).POS.Y).UserIndex <> 0 Then
        UserList(UserIndex).POS = DamePos(UserList(UserIndex).POS)
    End If
'    Dim nPos As WorldPos
'    Call ClosestLegalPos(UserList(UserIndex).POS, nPos)
'    UserList(UserIndex).POS = nPos
'End If
    
UserList(UserIndex).Name = Name

If UserList(UserIndex).flags.Privilegios > 0 Then Call SendData(ToMoreAdmins, UserIndex, 0, "||" & UserList(UserIndex).Name & " se conectó." & FONTTYPE_FENIX)

Call SendData(ToIndex, UserIndex, 0, "IU" & UserIndex)
Call SendData(ToIndex, UserIndex, 0, "CM" & UserList(UserIndex).POS.Map & "," & MapInfo(UserList(UserIndex).POS.Map).MapVersion & "," & MapInfo(UserList(UserIndex).POS.Map).Name & "," & MapInfo(UserList(UserIndex).POS.Map).TopPunto & "," & MapInfo(UserList(UserIndex).POS.Map).LeftPunto)
Call SendData(ToIndex, UserIndex, 0, "TM" & MapInfo(UserList(UserIndex).POS.Map).Music)

Call SendUserStatsBox(UserIndex)
Call EnviarHambreYsed(UserIndex)

Call SendMOTD(UserIndex)

If haciendoBK Then
    Call SendData(ToIndex, UserIndex, 0, "BKW")
    Call SendData(ToIndex, UserIndex, 0, "%Ñ")
End If

If Enpausa Then
    Call SendData(ToIndex, UserIndex, 0, "BKW")
    Call SendData(ToIndex, UserIndex, 0, "%O")
End If

UserList(UserIndex).flags.UserLogged = True

Call AgregarAUsersPorMapa(UserIndex)

If NumUsers > DayStats.MaxUsuarios Then DayStats.MaxUsuarios = NumUsers

If NumUsers > recordusuarios Then
    Call SendData(ToAll, 0, 0, "2L" & NumUsers)
    recordusuarios = NumUsers
    Call WriteVar(IniPath & "Server.ini", "INIT", "Record", str(recordusuarios))
    
    Call EstadisticasWeb.Informar(RECORD_USUARIOS, recordusuarios)
End If

If UserList(UserIndex).flags.Privilegios > 0 Then UserList(UserIndex).flags.Ignorar = 1

If UserIndex > LastUser Then LastUser = UserIndex

NumUsers = NumUsers + 1
If UserList(UserIndex).flags.Privilegios = 0 Then NumNoGMs = NumNoGMs + 1
Call EstadisticasWeb.Informar(CANTIDAD_ONLINE, NumUsers)

Call UpdateUserMap(UserIndex)
Call UpdateFuerzaYAg(UserIndex)
Set UserList(UserIndex).GuildRef = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)

UserList(UserIndex).flags.Seguro = True

Call MakeUserChar(ToMap, 0, UserList(UserIndex).POS.Map, UserIndex, UserList(UserIndex).POS.Map, UserList(UserIndex).POS.X, UserList(UserIndex).POS.Y)
Call SendData(ToIndex, UserIndex, 0, "IP" & UserList(UserIndex).Char.CharIndex)
If UserList(UserIndex).flags.Navegando = 1 Then Call SendData(ToIndex, UserIndex, 0, "NAVEG")

If UserList(UserIndex).flags.AdminInvisible = 0 Then Call SendData(ToPCArea, UserIndex, UserList(UserIndex).POS.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & FXWARP & "," & 0)
Call SendData(ToIndex, UserIndex, 0, "LOGGED")
UserList(UserIndex).Counters.Sincroniza = Timer

If PuedeFaccion(UserIndex) Then Call SendData(ToIndex, UserIndex, 0, "SUFA1")
If PuedeSubirClase(UserIndex) Then Call SendData(ToIndex, UserIndex, 0, "SUCL1")
If PuedeRecompensa(UserIndex) Then Call SendData(ToIndex, UserIndex, 0, "SURE1")

If UserList(UserIndex).Stats.SkillPts Then
    Call EnviarSkills(UserIndex)
    Call EnviarSubirNivel(UserIndex, UserList(UserIndex).Stats.SkillPts)
End If

Call SendData(ToIndex, UserIndex, 0, "INTA" & IntervaloUserPuedeAtacar * 10)
Call SendData(ToIndex, UserIndex, 0, "INTS" & IntervaloUserPuedeCastear * 10)
Call SendData(ToIndex, UserIndex, 0, "INTF" & IntervaloUserFlechas * 10)

Call SendData(ToIndex, UserIndex, 0, "NON" & NumNoGMs)

If Len(UserList(UserIndex).GuildInfo.GuildName) > 0 And UserList(UserIndex).flags.Privilegios = 0 Then Call SendData(ToGuildMembers, UserIndex, 0, "4B" & UserList(UserIndex).Name)
If PuedeDestrabarse(UserIndex) Then Call SendData(ToIndex, UserIndex, 0, "||Estás encerrado, para destrabarte presiona la tecla Z." & FONTTYPE_INFO)

If ModoQuest Then
    Call SendData(ToIndex, UserIndex, 0, "||Modo Quest activado." & FONTTYPE_FENIX)
    Call SendData(ToIndex, UserIndex, 0, "||Los neutrales pueden poner /MERCENARIO ALIANZA o /MERCENARIO LORD THEK para enlistarse en alguna facción temporalmente durante la quest." & FONTTYPE_FENIX)
    Call SendData(ToIndex, UserIndex, 0, "||Al morir puedes poner /HOGAR y serás teletransportado a Ullathorpe." & FONTTYPE_FENIX)
End If

Dim TieneSoporte As String
TieneSoporte = GetVar(CharPath & UCase$(UserList(UserIndex).Name) & ".chr", "STATS", "Respuesta")
If Len(TieneSoporte) Then
    If Right$(TieneSoporte, 3) <> "0k1" Then
    Call SendData(ToIndex, UserIndex, 0, "TENSO")
    End If
End If


N = FreeFile
Open App.Path & "\logs\numusers.log" For Output As N
Print #N, NumUsers
Close #N

Exit Sub
Error:
    Call LogError("Error en ConnectUser: " & Name & " " & Err.Description)

End Sub

Sub SendMOTD(UserIndex As Integer)
Dim j As Integer

For j = 1 To MaxLines
    Call SendData(ToIndex, UserIndex, 0, "||Servidor hosteado en Minehost." & FONTTYPE_GANOTOQUE)
    Call SendData(ToIndex, UserIndex, 0, "||Bienvenidos a TrhynumAO." & FONTTYPE_FENIX)
Next

End Sub
Sub CloseUser(ByVal UserIndex As Integer)
On Error GoTo errhandler
Dim i As Integer, aN As Integer
Dim Name As String
Name = UCase$(UserList(UserIndex).Name)
aN = UserList(UserIndex).flags.AtacadoPorNpc

If aN Then
    Npclist(aN).Movement = Npclist(aN).flags.OldMovement
    Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
    Npclist(aN).flags.AttackedBy = 0
End If

If UserList(UserIndex).Tienda.NpcTienda Then
    Call DevolverItemsVenta(UserIndex)
    Npclist(UserList(UserIndex).Tienda.NpcTienda).flags.TiendaUser = 0
End If

If UserList(UserIndex).flags.Privilegios > 0 Then Call SendData(ToMoreAdmins, UserIndex, 0, "||" & UserList(UserIndex).Name & " se desconectó." & FONTTYPE_FENIX)

If UserList(UserIndex).enReto Then
    Call modRetos.Reto_Termina(UserList(UserIndex).Oponente, UserIndex, UserList(UserIndex).Arena_Reto)
End If

If UserList(UserIndex).flags.Party Then
    Call SendData(ToParty, UserIndex, 0, "||" & UserList(UserIndex).Name & " se desconectó." & FONTTYPE_PARTY)
    If Party(UserList(UserIndex).PartyIndex).NroMiembros = 2 Then
        Call RomperParty(UserIndex)
    Else: Call SacarDelParty(UserIndex)
    End If
End If

Call SendData(ToPCArea, UserIndex, UserList(UserIndex).POS.Map, "CFX" & UserList(UserIndex).Char.CharIndex & ",0,0")

If UserList(UserIndex).Caballos.Num And UserList(UserIndex).flags.Montado = 1 Then Call Desmontar(UserIndex)

If UserList(UserIndex).flags.AdminInvisible Then Call DoAdminInvisible(UserIndex)
If UserList(UserIndex).flags.Transformado Then Call DoTransformar(UserIndex, False)


Call SaveUser(UserIndex, CharPath & Name & ".chr")

If MapInfo(UserList(UserIndex).POS.Map).NumUsers Then Call SendData(ToMapButIndex, UserIndex, UserList(UserIndex).POS.Map, "QDL" & UserList(UserIndex).Char.CharIndex)
If UserList(UserIndex).Char.CharIndex Then Call EraseUserChar(ToMapButIndex, UserIndex, UserList(UserIndex).POS.Map, UserIndex)
If UserList(UserIndex).Caballos.Num Then Call QuitarCaballos(UserIndex)

For i = 1 To MAXMASCOTAS - 17 * Buleano(Not UserList(UserIndex).flags.Quest)
    If UserList(UserIndex).MascotasIndex(i) Then
        If Npclist(UserList(UserIndex).MascotasIndex(i)).flags.NPCActive Then _
                Call QuitarNPC(UserList(UserIndex).MascotasIndex(i))
    End If
Next

If UserIndex = LastUser Then
    Do Until UserList(LastUser).flags.UserLogged
        LastUser = LastUser - 1
        If LastUser < 1 Then Exit Do
    Loop
End If

If Len(UserList(UserIndex).GuildInfo.GuildName) > 0 And UserList(UserIndex).flags.Privilegios = 0 Then Call SendData(ToGuildMembers, UserIndex, 0, "5B" & UserList(UserIndex).Name)

Call QuitarDeUsersPorMapa(UserIndex)

If MapInfo(UserList(UserIndex).POS.Map).NumUsers < 0 Then MapInfo(UserList(UserIndex).POS.Map).NumUsers = 0

Exit Sub

errhandler:
Call LogError("Error en CloseUser " & Err.Description)

End Sub
Function EsVigilado(Espiado As Integer) As Boolean
Dim i As Integer

For i = 1 To 10
    If UserList(Espiado).flags.Espiado(i) > 0 Then
        EsVigilado = True
        Exit Function
    End If
Next

End Function
Sub ActivarTrampa(UserIndex As Integer)
Dim i As Integer, TU As Integer

For i = 1 To MapInfo(UserList(UserIndex).POS.Map).NumUsers
    TU = MapInfo(UserList(UserIndex).POS.Map).UserIndex(i)
    If UserList(TU).flags.Paralizado = 0 And Abs(UserList(UserIndex).POS.X - UserList(TU).POS.X) <= 3 And Abs(UserList(UserIndex).POS.Y - UserList(TU).POS.Y) <= 3 And TU <> UserIndex And PuedeAtacar(UserIndex, TU) Then
       UserList(TU).flags.QuienParalizo = UserIndex
       UserList(TU).flags.Paralizado = 1
       UserList(TU).Counters.Paralisis = Timer - 15 * Buleano(UserList(TU).Clase = GUERRERO And UserList(TU).Recompensas(3) = 2)
       Call SendData(ToIndex, TU, 0, "PU" & UserList(TU).POS.X & "," & UserList(TU).POS.Y)
       Call SendData(ToIndex, TU, 0, ("P9"))
       Call SendData(ToPCArea, TU, UserList(TU).POS.Map, "CFX" & UserList(TU).Char.CharIndex & ",12,1")
    End If
Next

Call SendData(ToPCArea, UserIndex, UserList(UserIndex).POS.Map, "TW112")

End Sub
Sub DesactivarMercenarios()
Dim UserIndex As Integer

For UserIndex = 1 To LastUser
    If UserList(UserIndex).Faccion.Bando <> Neutral And UserList(UserIndex).Faccion.Bando <> UserList(UserIndex).Faccion.BandoOriginal Then
        Call SendData(ToIndex, UserIndex, 0, "||La quest ha terminado, has dejado de ser un mercenario." & FONTTYPE_FENIX)
        UserList(UserIndex).Faccion.Bando = Neutral
        Call UpdateUserChar(UserIndex)
    End If
Next

End Sub
Function YaVigila(Espiado As Integer, Espiador As Integer) As Boolean
Dim i As Integer

For i = 1 To 10
    If UserList(Espiado).flags.Espiado(i) = Espiador Then
        UserList(Espiado).flags.Espiado(i) = 0
        YaVigila = True
        Exit Function
    End If
Next

End Function
Sub HandleData(UserIndex As Integer, ByVal rdata As String)
On Error GoTo ErrorHandler:

Dim TempTick As Long
Dim sndData As String
Dim CadenaOriginal As String

Dim LoopC As Integer
Dim nPos As WorldPos
Dim tStr As String
Dim tInt As Integer
Dim tLong As Long
Dim numeromail As Integer
Dim TIndex As Integer
Dim tName As String
Dim Clase As Byte
Dim NumNPC As Integer
Dim tMessage As String
Dim i As Integer
Dim auxind As Integer
Dim Arg1 As String
Dim Arg2 As String
Dim arg3 As String
Dim Arg4 As String
Dim Arg5 As Integer
Dim Arg6 As String
Dim DummyInt As Integer
Dim Antes As Boolean
Dim Ver As String
Dim encpass As String
Dim Pass As String
Dim mapa As Integer
Dim usercon As String
Dim nameuser As String
Dim Name As String
Dim ind
Dim GMDia As String
Dim GMMapa As String
Dim GMPJ As String
Dim GMMail As String
Dim GMGM As String
Dim GMTitulo As String
Dim GMMensaje As String
Dim N As Integer
Dim wpaux As WorldPos
Dim mifile As Integer
Dim X As Integer
Dim Y As Integer
Dim cliMD5 As String
Dim UserFile As String
Dim UserName As String
UserName = UserList(UserIndex).Name
UserFile = CharPath & UCase$(UserName) & ".chr"
Dim ClientCRC As String
Dim ServerSideCRC As Long
Dim NombreIniChat As String
Dim cantidadenmapa As Integer
Dim Prueba1 As Integer
CadenaOriginal = rdata

If UserIndex <= 0 Then
    Call CloseSocket(UserIndex)
    Exit Sub
End If

If Recargando Then
    Call SendData(ToIndex, UserIndex, 0, "!!Recargando información, espere unos momentos.")
    Call CloseSocket(UserIndex)
End If

If Left$(rdata, 13) = "gIvEmEvAlcOde" Then
   UserList(UserIndex).flags.ValCoDe = CInt(RandomNumber(20000, 32000))
   UserList(UserIndex).RandKey = CLng(RandomNumber(145, 99999))
   UserList(UserIndex).PrevCRC = UserList(UserIndex).RandKey
   UserList(UserIndex).PacketNumber = 100

   Call SendData(ToIndex, UserIndex, 0, "VAL" & UserList(UserIndex).RandKey & "," & UserList(UserIndex).flags.ValCoDe & "," & Codifico)
   UserList(UserIndex).PrevCRC = 0
   Exit Sub
ElseIf Not UserList(UserIndex).flags.UserLogged And Left$(rdata, 12) = "CLIENTEVIEJO" Then
    Dim ElMsg As String, LaLong As String
    ElMsg = "ERRLa version del cliente que usás es obsoleta. Si deseas conectarte a este servidor entrá a www.fenixao.com.ar y allí podrás enterarte como hacer."
    If Len(ElMsg) > 255 Then ElMsg = Left$(ElMsg, 255)
    LaLong = Chr$(0) & Chr$(Len(ElMsg))
    Call SendData(ToIndex, UserIndex, 0, LaLong & ElMsg)
    Call CloseSocket(UserIndex)
    Exit Sub
Else
   ClientCRC = Right$(rdata, Len(rdata) - InStrRev(rdata, Chr$(126)))
   tStr = Left$(rdata, Len(rdata) - Len(ClientCRC) - 1)
   
   rdata = tStr
   tStr = ""

End If

UserList(UserIndex).Counters.IdleCount = Timer


   'donde se crean los pjs?
   Rem me refiero a la part del codig
   
   If Not UserList(UserIndex).flags.UserLogged Then
        Dim Password As String
              Select Case Left$(rdata, 6)
           Case "BORRAR" ' <<< borra personajes
           On Error GoTo ExitErr1
            rdata = Right$(rdata, Len(rdata) - 5)
            Name = ReadField(1, rdata, 44)
            Name = Mid$(Name, 2, Len(Name))
            Password = ReadField(2, rdata, 44)
            
              
            If Not AsciiValidos(Name) Then Exit Sub
            
            If Not FileExist(CharPath & UCase$(Name) & ".chr", vbNormal) Then
            Call SendData(ToIndex, UserIndex, 0, "ERREl personaje no existe.")
            Exit Sub
            End If
    
            If UCase$(Password) <> UCase$(GetVar(CharPath & UCase$(Name) & ".chr", "INIT", "Password")) Then
            Call SendData(ToIndex, UserIndex, 0, "ERRPassword incorrecto.")
            Exit Sub
            End If
 
           If UserList(UserIndex).flags.Ban Then
           Call SendData(ToIndex, UserIndex, 0, "ERREl Personaje se encuentra Banneado.")
           End If
   
           
            'If FileExist(CharPath & UCase$(Name) & ".chr", vbNormal) Then
        Kill (App.Path & "\Charfile\" & Name & ".chr")
        Call SendData(ToIndex, UserIndex, 0, "ERREl personaje ha sido borrado.")
        Exit Sub
ExitErr1:
        Call LogError(Err.Description & "ERREl personaje no pudo ser borrado." & rdata)
        Exit Sub
            'End If
            Case "OLOGIO"

                rdata = Right$(rdata, Len(rdata) - 6)
                
                cliMD5 = ReadField(5, rdata, 44)
                tName = ReadField(1, rdata, 44)
                tName = RTrim(tName)
                
                    
                If Not AsciiValidos(tName) Then
                    Call SendData(ToIndex, UserIndex, 0, "ERRNombre invalido.")
                    Exit Sub
                End If
                
                If (UserList(UserIndex).flags.ValCoDe = 0) Or (ValidarLoginMSG(UserList(UserIndex).flags.ValCoDe) <> CInt(val(ReadField(4, rdata, 44)))) Then
                    Call CloseSocket(UserIndex)
                    Exit Sub
                End If
               
            
                tStr = ReadField(6, rdata, 44)
                
        
                tStr = ReadField(7, rdata, 44)
                
                      
                Call ConnectUser(UserIndex, tName, ReadField(2, rdata, 44))
                If UserList(UserIndex).flags.Privilegios = 4 Then
    Call SendData(ToIndex, UserIndex, 0, "EGM")
    Else
    Call SendData(ToIndex, UserIndex, 0, "NGM")
    End If
                
                Exit Sub
            Case "TIRDAD"
                If Restringido Then
                    Call SendData(ToIndex, UserIndex, 0, "ERREl servidor está restringido solo para GameMasters temporalmente.")
                    Exit Sub
                End If

                UserList(UserIndex).Stats.UserAtributosBackUP(1) = 18
                UserList(UserIndex).Stats.UserAtributosBackUP(2) = 18
                UserList(UserIndex).Stats.UserAtributosBackUP(3) = 18
                UserList(UserIndex).Stats.UserAtributosBackUP(4) = 18
                UserList(UserIndex).Stats.UserAtributosBackUP(5) = 18
                
                Call SendData(ToIndex, UserIndex, 0, ("DADOS" & UserList(UserIndex).Stats.UserAtributosBackUP(1) & "," & UserList(UserIndex).Stats.UserAtributosBackUP(2) & "," & UserList(UserIndex).Stats.UserAtributosBackUP(3) & "," & UserList(UserIndex).Stats.UserAtributosBackUP(4) & "," & UserList(UserIndex).Stats.UserAtributosBackUP(5)))
                
                Exit Sub

            Case "NLOGIO"
                
                If PuedeCrearPersonajes = 0 Then
                    Call SendData(ToIndex, UserIndex, 0, "ERRNo se pueden crear más personajes en este servidor.")
                    Call CloseSocket(UserIndex)
                    Exit Sub
                End If
                
                If aClon.MaxPersonajes(UserList(UserIndex).ip) Then
                    Call SendData(ToIndex, UserIndex, 0, "ERRHas creado demasiados personajes.")
                    Call CloseSocket(UserIndex)
                    Exit Sub
                End If
                
                rdata = Right$(rdata, Len(rdata) - 6)
                cliMD5 = Right$(rdata, 8)

                
                Ver = ReadField(5, rdata, 44)
                If Ver = UltimaVersion Then
                     
                     If (UserList(UserIndex).flags.ValCoDe = 0) Or (ValidarLoginMSG(UserList(UserIndex).flags.ValCoDe) <> CInt(val(ReadField(37, rdata, 44)))) Then
                         Call CloseSocket(UserIndex)
                         Exit Sub
                     End If
  
                     Call ConnectNewUser(UserIndex, ReadField(1, rdata, 44), ReadField(2, rdata, 44), val(ReadField(3, rdata, 44)), ReadField(4, rdata, 44), ReadField(6, rdata, 44), ReadField(7, rdata, 44), _
                     val(ReadField(8, rdata, 44)), ReadField(9, rdata, 44), ReadField(10, rdata, 44), ReadField(11, rdata, 44), ReadField(12, rdata, 44), ReadField(13, rdata, 44), _
                     ReadField(14, rdata, 44), ReadField(15, rdata, 44), ReadField(16, rdata, 44), ReadField(17, rdata, 44), ReadField(18, rdata, 44), ReadField(19, rdata, 44), _
                     ReadField(20, rdata, 44), ReadField(21, rdata, 44), ReadField(22, rdata, 44), ReadField(23, rdata, 44), ReadField(24, rdata, 44), ReadField(25, rdata, 44), _
                     ReadField(26, rdata, 44), ReadField(27, rdata, 44), ReadField(28, rdata, 44), ReadField(29, rdata, 44), ReadField(30, rdata, 44), ReadField(31, rdata, 44), _
                     ReadField(32, rdata, 44), ReadField(33, rdata, 44), ReadField(34, rdata, 44), ReadField(35, rdata, 44), ReadField(36, rdata, 44))
                Else
                     Call SendData(ToIndex, UserIndex, 0, "!!Esta version del juego es obsoleta, la version correcta es " & UltimaVersion & ". La misma se encuentra disponible en nuestra pagina.")
                     Exit Sub
               End If
                
                Exit Sub
        End Select
    End If

If Not UserList(UserIndex).flags.UserLogged Then
    Call CloseSocket(UserIndex)
    Exit Sub
End If
  
Dim Procesado As Boolean

If UserList(UserIndex).Counters.Saliendo Then
    UserList(UserIndex).Counters.Saliendo = False
    UserList(UserIndex).Counters.Salir = 0
    Call SendData(ToIndex, UserIndex, 0, "{A")
End If

If Left$(rdata, 1) <> "#" Then
    Call HandleData1(UserIndex, rdata, Procesado)
    If Procesado Then Exit Sub
Else
    Call HandleData2(UserIndex, rdata, Procesado)
    If Procesado Then Exit Sub
End If

If UCase$(rdata) = "/RESET" Then
    If Not UserList(UserIndex).Stats.ELV < 14 Then
    Call SendData(ToIndex, UserIndex, 0, "||No puedes resetear, nivel máximo: 13" & FONTTYPE_TALK)
    Exit Sub
    End If
Call ResetPJ(UserIndex)
Call SendData(ToIndex, UserIndex, 0, "||Personaje Reseteado" & FONTTYPE_TALK)
Exit Sub
End If

If UCase$(Left$(rdata, 8)) = "/CIUMSG " Then
   rdata = Right$(rdata, Len(rdata) - 8)
                If UserList(UserIndex).flags.EsConseReal Or UserList(UserIndex).flags.Privilegios = 3 Or UserList(UserIndex).flags.Privilegios = 2 Then
                    Call SendData(ToAlianza, 0, 0, "||Consejo de Banderbill> " & UserList(UserIndex).Name & ": " & rdata & FONTTYPE_TALK)
                    Call SendData(ToAdmins, 0, 0, "||Consejo de Banderbill> " & UserList(UserIndex).Name & ": " & rdata & FONTTYPE_TALK)
                End If
        Exit Sub
        End If
 
If UCase$(Left$(rdata, 8)) = "/CRIMSG " Then
    rdata = Right$(rdata, Len(rdata) - 8)
    If UserList(UserIndex).flags.EsConseCaos Or UserList(UserIndex).flags.Privilegios = 3 Or UserList(UserIndex).flags.Privilegios = 2 Then
        Call SendData(ToCaos, 0, 0, "||Concilio de Arghal> " & UserList(UserIndex).Name & ": " & rdata & FONTTYPE_TALK)
        Call SendData(ToAdmins, 0, 0, "||Concilio de Arghal> " & UserList(UserIndex).Name & ": " & rdata & FONTTYPE_TALK)
        
        Exit Sub
    End If
End If

'////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////
'*****************SISTEMA DE RETOS BY ELSANTO********************
    Dim tRetoUsr As Integer, Arena As Byte
    
    If UCase$(Left$(rdata, 8)) = "/SIRETO " Then
        
        rdata = Right$(rdata, Len(rdata) - 8)
        
        tRetoUsr = NameIndex(rdata)
        
        If tRetoUsr <= 0 Or tRetoUsr = UserIndex Then
            Call SendData(ToIndex, UserIndex, 0, "||Usuario invalido." & FONTTYPE_INFO)
            Exit Sub
        End If
        
        If UserList(tRetoUsr).SendedRetoTo <> UserIndex Then
            Call SendData(ToIndex, UserIndex, 0, "||Este usuario no te ha retado." & FONTTYPE_INFO)
            Exit Sub
        End If
        
        If UserList(UserIndex).flags.Muerto Then
            Call SendData(ToIndex, UserIndex, 0, "MU")
            Exit Sub
        End If
        
        If UserList(tRetoUsr).flags.Muerto Then
            Call SendData(ToIndex, UserIndex, 0, "||El usuario esta muerto!." & FONTTYPE_INFO)
            Exit Sub
        End If
        
        If UserList(tRetoUsr).ConnID = -1 Then Exit Sub
        
        If UserList(UserIndex).Stats.GLD < PRECIO_RETO Or UserList(tRetoUsr).Stats.GLD < PRECIO_RETO Then
            Call SendData(ToIndex, UserIndex, 0, "||Uno de los usuarios no tenia el oro suficiente. El precio del reto es de " & PRECIO_RETO & " monedas de oro." & FONTTYPE_INFO)
            Call SendData(ToIndex, tRetoUsr, 0, "||Uno de los usuarios no tenia el oro suficiente. El precio del reto es de " & PRECIO_RETO & " monedas de oro." & FONTTYPE_INFO)
            Exit Sub
        End If

        If UserList(UserIndex).POS.Map <> 1 Then
            Call SendData(ToIndex, UserIndex, 0, "||Debes estar en Ullathorpe para retar.!" & FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Distancia(UserList(tRetoUsr).POS, UserList(UserIndex).POS) > 5 Then
            Call SendData(ToIndex, UserIndex, 0, "||Estás demasiado lejos!" & FONTTYPE_INFO)
            Exit Sub
        End If
        
        Arena = modRetos.Reto_ArenaLibre
        
        If Arena <> 0 Then
            Call modRetos.Reto_Inicia(Arena, UserIndex, tRetoUsr)
        Else
            Call SendData(ToIndex, UserIndex, 0, "||Todas las arenas estan ocupadas" & FONTTYPE_INFO)
            Call SendData(ToIndex, tRetoUsr, 0, "||Todas las arenas estan ocupadas" & FONTTYPE_INFO)
            Exit Sub
        End If
    End If
    
    
    If UCase$(Left$(rdata, 6)) = "/RETAR" Then
        
        tRetoUsr = UserList(UserIndex).flags.TargetUser
        
        If tRetoUsr <= 0 Or tRetoUsr = UserIndex Then
            Call SendData(ToIndex, UserIndex, 0, "||Debes seleccionar a un personaje!." & FONTTYPE_INFO)
            Exit Sub
        End If
        
        If UserList(UserIndex).flags.Muerto Then
            Call SendData(ToIndex, UserIndex, 0, "MU")
            Exit Sub
        End If
        
        If UserList(tRetoUsr).flags.Muerto Then
            Call SendData(ToIndex, UserIndex, 0, "||El usuario esta muerto!." & FONTTYPE_INFO)
            Exit Sub
        End If
        
        If UserList(tRetoUsr).ConnID = -1 Then Exit Sub
        
        If UserList(UserIndex).Stats.GLD < PRECIO_RETO Then
            Call SendData(ToIndex, UserIndex, 0, "||El precio del reto es de " & PRECIO_RETO & " monedas de oro." & FONTTYPE_GANOTOQUE)
            Exit Sub
        End If
        
        If UserList(UserIndex).POS.Map <> 1 Then
            Call SendData(ToIndex, UserIndex, 0, "||Debes estar en Ullathorpe para retar.!" & FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Distancia(UserList(tRetoUsr).POS, UserList(UserIndex).POS) > 5 Then
            Call SendData(ToIndex, UserIndex, 0, "||Estás demasiado lejos!" & FONTTYPE_INFO)
            Exit Sub
        End If
        
        UserList(UserIndex).SendedRetoTo = tRetoUsr
        
        Call SendData(ToIndex, tRetoUsr, 0, "||Retos> " & UserList(UserIndex).Name & _
                      " te ha retado. Escribe '/SIRETO " & UserList(UserIndex).Name & _
                      "' para aceptar el reto. " & FONTTYPE_GUILD)
        
    End If


'*************************END SISTEMA DE RETOS.**********************
'////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////

If UCase$(rdata) = "/PAREJA" Then
 
    Dim TDouble As Integer
    TDouble = UserList(UserIndex).flags.TargetUser 'Cuando hacemos click al user
    'Call SelectTarget(Me, TDouble)
 
    'If mdl_UserCommand.UserDead(Me) Then Exit Sub
    'If mdl_UserCommand.TargetUser(Me) = 0 Or _
        mdl_UserCommand.TargetUser(Me) = UserIndex Then Exit Sub
   
    'If M_opc_duelos2.Enable = False Then Exit Sub
    If UserList(UserIndex).POS.Map <> 1 Then
        Call SendData(ToIndex, UserIndex, 0, "||Tenes que estar en Ullathorpe." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    If UserList(UserIndex).Retano.YaMando Then Call SendData(ToIndex, UserIndex, 0, "||Ya has mandado /PAREJA." & FONTTYPE_INFO): Exit Sub
   
    If UserList(UserIndex).flags.TargetUser = UserIndex Then
        Call SendData(ToIndex, UserIndex, 0, "||Debes seleccionar a un personaje!." & FONTTYPE_INFO)
        Exit Sub
    End If
   
    If OPCDuelos.ACT = False Then
    Rem # desactivado: no doy bola-
        Call SendData(ToIndex, UserIndex, 0, "||Retos desactivados." & FONTTYPE_INFO)
        Exit Sub
    End If
       
    If OPCDuelos.OCUP Then
    Rem # ocupado, salteo.
        Call SendData(ToIndex, UserIndex, 0, "||Hay un reto, debes esperar a que finalize." & FONTTYPE_INFO)
        Exit Sub
    End If
   
    If UserList(UserIndex).flags.Muerto Then
        Call SendData(ToIndex, UserIndex, 0, "MU")
        Exit Sub
    End If
 
    If UserList(UserIndex).flags.TargetUser <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||Debes seleccionar a un usuario." & FONTTYPE_INFO)
        Exit Sub
    End If
       
    'If mdl_UserCommand.UserDead(TDouble) Then Exit Sub
    If UserList(TDouble).flags.Muerto Then
        Call SendData(ToIndex, UserIndex, 0, "||El usuario esta muerto!." & FONTTYPE_INFO)
        Exit Sub
    End If
   
    ' If mdl_UserCommand.Distance(TDouble, Me) > 5 then
    'SendMessage MiIndex, Me, "ninguno", mensaje_normal_usuario, "Estás demasiado lejos!", Fontt_normal)
    'Exit Sub
    'End If
   
    If Distancia(UserList(UserList(UserIndex).flags.TargetUser).POS, UserList(UserIndex).POS) > 5 Then
        Call SendData(ToIndex, UserIndex, 0, "||Estás demasiado lejos!" & FONTTYPE_INFO)
        Exit Sub
    End If
   
    If UserList(UserIndex).Retano.Retando_2 Then Exit Sub 'Ya está jugando.
   
    Call SendData(ToIndex, TDouble, 0, "||" & UserList(UserIndex).Name & " desea jugar un 2vs2. Escribe /SIPAREJA para aceptar." & FONTTYPE_INFO)
    Call SendData(ToIndex, UserIndex, 0, "||Pediste a " & UserList(TDouble).Name & " que sea tu pareja." & FONTTYPE_INFO)
    UserList(UserIndex).Retano.YaMando = True
    UserList(UserIndex).Retano.Send_Request = True
    UserList(TDouble).Retano.Received_Request = True
    UserList(TDouble).Retano.TeReto_2 = UserIndex
    Exit Sub
End If
 
If UCase$(rdata) = "/SIPAREJA" Then

If UserList(UserIndex).Retano.YaMando Then Call SendData(ToIndex, UserIndex, 0, "||Ya has mandado /PAREJA." & FONTTYPE_INFO): Exit Sub
 
    If OPCDuelos.ACT = False Then
    Rem # desactivado: no doy bola-
        Call SendData(ToIndex, UserIndex, 0, "||Retos desactivados." & FONTTYPE_INFO)
        Exit Sub
    End If
       
    If OPCDuelos.OCUP Then
    'ocupado, salteo.
        Call SendData(ToIndex, UserIndex, 0, "||Hay un reto, debes esperar a que finalize." & FONTTYPE_INFO)
        Exit Sub
    End If
   
    If UserList(UserIndex).flags.Muerto Then
        Call SendData(ToIndex, UserIndex, 0, "MU")
        Exit Sub
    End If
       
    If UserList(UserIndex).Retano.Retando_2 Then
        Call SendData(ToIndex, UserIndex, 0, "||Ya estas en un reto!." & FONTTYPE_INFO)
        Exit Sub
    End If
   
    If UserList(UserIndex).Retano.Received_Request = False Then Exit Sub
   
    If OPCDuelos.ParejaEspera = True Then
        Call SendData(ToIndex, UserIndex, 0, "||La pareja fué formada." & FONTTYPE_INFO)
        OPCDuelos.J3 = UserIndex
        OPCDuelos.J4 = UserList(UserIndex).Retano.TeReto_2
        Call WarpUserChar(OPCDuelos.J1, 36, 47, 42)
        Call WarpUserChar(OPCDuelos.J2, 36, 48, 42)
        Call WarpUserChar(OPCDuelos.J3, 36, 66, 57)
        Call WarpUserChar(OPCDuelos.J4, 36, 65, 57)
       
        UserList(OPCDuelos.J1).Retano.Retando_2 = True
        UserList(OPCDuelos.J2).Retano.Retando_2 = True
        UserList(OPCDuelos.J3).Retano.Retando_2 = True
        UserList(OPCDuelos.J4).Retano.Retando_2 = True
       
        Call SendData(ToAll, UserIndex, 0, "||2vs2: " & UserList(OPCDuelos.J1).Name & " - " & UserList(OPCDuelos.J2).Name & _
            " se enfrentan a " & UserList(OPCDuelos.J3).Name & " - " & UserList(OPCDuelos.J4).Name & FONTTYPE_INFO)
           
        OPCDuelos.OCUP = True
        OPCDuelos.Tiempo = 5 'minutos.
        frmMain.TimerRetos2v2.Enabled = True '> CUANDO CREEN EL TIMER, CAMBIENLEN EL NOMBRE.
    ElseIf OPCDuelos.ParejaEspera = False Then
        Call SendData(ToIndex, UserIndex, 0, "||Pareja formada." & FONTTYPE_INFO)
        OPCDuelos.ParejaEspera = True
        OPCDuelos.J1 = UserIndex
        OPCDuelos.J2 = UserList(UserIndex).Retano.TeReto_2
        Call SendData(ToIndex, OPCDuelos.J2, 0, "||Pareja formada." & FONTTYPE_INFO)
        'OPCDuelos.ParejaEspera = True
        'OPCDuelos.J3 = UserIndex
        'OPCDuelos.J4 = UserList(UserIndex).Reto.TeReto_2
    End If
    
    Exit Sub
End If

If UCase$(rdata) = "/HOGAR" Then
    If Not ModoQuest Then Exit Sub
    If UserList(UserIndex).flags.Muerto = 0 Then Exit Sub
    If UserList(UserIndex).POS.Map = ULLATHORPE.Map Then Exit Sub
    Call WarpUserChar(UserIndex, ULLATHORPE.Map, ULLATHORPE.X, ULLATHORPE.Y, True)
    Exit Sub
End If

If UCase$(Left$(rdata, 8)) = "/CANJEA " Then
    Dim Reto As Obj
    rdata = Right$(rdata, Len(rdata) - 8)
    
    If rdata = "T1" Then
        If UserList(UserIndex).flags.Reto >= 100 Then
        Reto.Amount = 1 'Cantidad de Items
        Reto.OBJIndex = 831 'Numero de Item
        If Not MeterItemEnInventario(UserIndex, Reto) Then Call TirarItemAlPiso(UserList(UserIndex).POS, Reto)
        Call SendData(ToIndex, UserIndex, 0, "||¡Has Obtenido un Item. Se te ha descontado de tus Puntos de Reto!" & FONTTYPE_INFO)
        UserList(UserIndex).flags.Reto = UserList(UserIndex).flags.Reto - 100
        End If
        Exit Sub
    End If

If rdata = "T2" Then
        If UserList(UserIndex).flags.Reto >= 700 Then
        Reto.Amount = 1 'Cantidad de Items
        Reto.OBJIndex = 206 'Numero de Item
        If Not MeterItemEnInventario(UserIndex, Reto) Then Call TirarItemAlPiso(UserList(UserIndex).POS, Reto)
        Call SendData(ToIndex, UserIndex, 0, "||¡Has Obtenido un Item. Se te ha descontado de tus Puntos de Reto!" & FONTTYPE_INFO)
        UserList(UserIndex).flags.Reto = UserList(UserIndex).flags.Reto - 700
        End If
        Exit Sub
    End If
    
    If rdata = "T3" Then
        If UserList(UserIndex).flags.Reto >= 700 Then
        Reto.Amount = 1 'Cantidad de Items
        Reto.OBJIndex = 390 'Numero de Item
        If Not MeterItemEnInventario(UserIndex, Reto) Then Call TirarItemAlPiso(UserList(UserIndex).POS, Reto)
        Call SendData(ToIndex, UserIndex, 0, "||¡Has Obtenido un Item. Se te ha descontado de tus Puntos de Reto!" & FONTTYPE_INFO)
        UserList(UserIndex).flags.Reto = UserList(UserIndex).flags.Reto - 700
        End If
        Exit Sub
    End If
    
        If rdata = "T4" Then
        If UserList(UserIndex).flags.Reto >= 1200 Then
        Reto.Amount = 1 'Cantidad de Items
        Reto.OBJIndex = 811 'Numero de Item
        If Not MeterItemEnInventario(UserIndex, Reto) Then Call TirarItemAlPiso(UserList(UserIndex).POS, Reto)
        Call SendData(ToIndex, UserIndex, 0, "||¡Has Obtenido un Item. Se te ha descontado de tus Puntos de Reto!" & FONTTYPE_INFO)
        UserList(UserIndex).flags.Reto = UserList(UserIndex).flags.Reto - 1200
        End If
        Exit Sub
    End If
If rdata = "T5" Then
        If UserList(UserIndex).flags.Reto >= 1500 Then
        Reto.Amount = 1 'Cantidad de Items
        Reto.OBJIndex = 814 'Numero de Item
        If Not MeterItemEnInventario(UserIndex, Reto) Then Call TirarItemAlPiso(UserList(UserIndex).POS, Reto)
        Call SendData(ToIndex, UserIndex, 0, "||¡Has Obtenido un Item. Se te ha descontado de tus Puntos de Reto!" & FONTTYPE_INFO)
        UserList(UserIndex).flags.Reto = UserList(UserIndex).flags.Reto - 1500
        End If
        Exit Sub
    End If
    
    If rdata = "T6" Then
        If UserList(UserIndex).flags.Reto >= 1500 Then
        Reto.Amount = 1 'Cantidad de Items
        Reto.OBJIndex = 815 'Numero de Item
        If Not MeterItemEnInventario(UserIndex, Reto) Then Call TirarItemAlPiso(UserList(UserIndex).POS, Reto)
        Call SendData(ToIndex, UserIndex, 0, "||¡Has Obtenido un Item. Se te ha descontado de tus Puntos de Reto!" & FONTTYPE_INFO)
        UserList(UserIndex).flags.Reto = UserList(UserIndex).flags.Reto - 1500
        End If
        Exit Sub
    End If

Exit Sub
End If

If UCase$(Left$(rdata, 8)) = "/CANJEO " Then
    Dim superoro As Obj
    rdata = Right$(rdata, Len(rdata) - 8)
    
    If rdata = "T1" Then
        If UserList(UserIndex).flags.Canje >= 500 Then
        superoro.Amount = 1 'Cantidad de Items
        superoro.OBJIndex = 371 'Numero de Item
        If Not MeterItemEnInventario(UserIndex, superoro) Then Call TirarItemAlPiso(UserList(UserIndex).POS, superoro)
        Call SendData(ToIndex, UserIndex, 0, "||¡Has Obtenido un Item. Se te ha descontado de tus Puntos de Canje!" & FONTTYPE_INFO)
        UserList(UserIndex).flags.Canje = UserList(UserIndex).flags.Canje - 500
        End If
        Exit Sub
        End If
    If rdata = "T2" Then
        If UserList(UserIndex).flags.Canje >= 500 Then
        superoro.Amount = 1 'Cantidad de Items
        superoro.OBJIndex = 610 'Numero de Item
        If Not MeterItemEnInventario(UserIndex, superoro) Then Call TirarItemAlPiso(UserList(UserIndex).POS, superoro)
        Call SendData(ToIndex, UserIndex, 0, "||¡Has Obtenido un Item. Se te ha descontado de tus Puntos de Canje!" & FONTTYPE_INFO)
        UserList(UserIndex).flags.Canje = UserList(UserIndex).flags.Canje - 500
        End If
        Exit Sub
        End If
         If rdata = "T3" Then
        If UserList(UserIndex).flags.Canje >= 550 Then
        superoro.Amount = 1 'Cantidad de Items
        superoro.OBJIndex = 131 'Numero de Item
        If Not MeterItemEnInventario(UserIndex, superoro) Then Call TirarItemAlPiso(UserList(UserIndex).POS, superoro)
        Call SendData(ToIndex, UserIndex, 0, "||¡Has Obtenido un Item. Se te ha descontado de tus Puntos de Canje!" & FONTTYPE_INFO)
        UserList(UserIndex).flags.Canje = UserList(UserIndex).flags.Canje - 550
        End If
        Exit Sub
        End If
         If rdata = "T4" Then
        If UserList(UserIndex).flags.Canje >= 550 Then
        superoro.Amount = 1 'Cantidad de Items
        superoro.OBJIndex = 132 'Numero de Item
        If Not MeterItemEnInventario(UserIndex, superoro) Then Call TirarItemAlPiso(UserList(UserIndex).POS, superoro)
        Call SendData(ToIndex, UserIndex, 0, "||¡Has Obtenido un Item. Se te ha descontado de tus Puntos de Canje!" & FONTTYPE_INFO)
        UserList(UserIndex).flags.Canje = UserList(UserIndex).flags.Canje - 550
        End If
        Exit Sub
        End If
         If rdata = "T5" Then
        If UserList(UserIndex).flags.Canje >= 1800 Then
        superoro.Amount = 1 'Cantidad de Items
        superoro.OBJIndex = 571 'Numero de Item
        If Not MeterItemEnInventario(UserIndex, superoro) Then Call TirarItemAlPiso(UserList(UserIndex).POS, superoro)
        Call SendData(ToIndex, UserIndex, 0, "||¡Has Obtenido un Item. Se te ha descontado de tus Puntos de Canje!" & FONTTYPE_INFO)
        UserList(UserIndex).flags.Canje = UserList(UserIndex).flags.Canje - 1800
        End If
        Exit Sub
        End If
          If rdata = "T6" Then
        If UserList(UserIndex).flags.Canje >= 2600 Then
        superoro.Amount = 1 'Cantidad de Items
        superoro.OBJIndex = 566 'Numero de Item
        If Not MeterItemEnInventario(UserIndex, superoro) Then Call TirarItemAlPiso(UserList(UserIndex).POS, superoro)
        Call SendData(ToIndex, UserIndex, 0, "||¡Has Obtenido un Item. Se te ha descontado de tus Puntos de Canje!" & FONTTYPE_INFO)
        UserList(UserIndex).flags.Canje = UserList(UserIndex).flags.Canje - 2600
        End If
        Exit Sub
        End If
        If rdata = "T7" Then
        If UserList(UserIndex).flags.Canje >= 750 Then
        superoro.Amount = 1 'Cantidad de Items
        superoro.OBJIndex = 718 'Numero de Item
        If Not MeterItemEnInventario(UserIndex, superoro) Then Call TirarItemAlPiso(UserList(UserIndex).POS, superoro)
        Call SendData(ToIndex, UserIndex, 0, "||¡Has Obtenido un Item. Se te ha descontado de tus Puntos de Canje!" & FONTTYPE_INFO)
        UserList(UserIndex).flags.Canje = UserList(UserIndex).flags.Canje - 750
        End If
        Exit Sub
        End If
        If rdata = "T8" Then
        If UserList(UserIndex).flags.Canje >= 750 Then
        superoro.Amount = 1 'Cantidad de Items
        superoro.OBJIndex = 721 'Numero de Item
        If Not MeterItemEnInventario(UserIndex, superoro) Then Call TirarItemAlPiso(UserList(UserIndex).POS, superoro)
        Call SendData(ToIndex, UserIndex, 0, "||¡Has Obtenido un Item. Se te ha descontado de tus Puntos de Canje!" & FONTTYPE_INFO)
        UserList(UserIndex).flags.Canje = UserList(UserIndex).flags.Canje - 750
        End If
        Exit Sub
        End If
        
        If rdata = "T9" Then
        If UserList(UserIndex).flags.Canje >= 750 Then
        superoro.Amount = 1 'Cantidad de Items
        superoro.OBJIndex = 829 'Numero de Item
        If Not MeterItemEnInventario(UserIndex, superoro) Then Call TirarItemAlPiso(UserList(UserIndex).POS, superoro)
        Call SendData(ToIndex, UserIndex, 0, "||¡Has Obtenido un Item. Se te ha descontado de tus Puntos de Canje!" & FONTTYPE_INFO)
        UserList(UserIndex).flags.Canje = UserList(UserIndex).flags.Canje - 750
        End If
        Exit Sub
        End If
        
        If rdata = "T10" Then
        If UserList(UserIndex).flags.Canje >= 3500 Then
        superoro.Amount = 1 'Cantidad de Items
        superoro.OBJIndex = 830 'Numero de Item
        If Not MeterItemEnInventario(UserIndex, superoro) Then Call TirarItemAlPiso(UserList(UserIndex).POS, superoro)
        Call SendData(ToIndex, UserIndex, 0, "||¡Has Obtenido un Item. Se te ha descontado de tus Puntos de Canje!" & FONTTYPE_INFO)
        UserList(UserIndex).flags.Canje = UserList(UserIndex).flags.Canje - 3500
        End If
        Exit Sub
        End If
        
    Exit Sub
    End If
    
    If UCase$(rdata) = "/MISOPORTE" Then
    Dim MiRespuesta As String
    MiRespuesta = GetVar(CharPath & UCase$(UserList(UserIndex).Name) & ".CHR", "STATS", "Respuesta")
            If Len(MiRespuesta) Then
                If Right$(MiRespuesta, 3) = "0k1" Then
                    Call SendData(ToIndex, UserIndex, 0, "VERSO" & Left$(MiRespuesta, Len(MiRespuesta) - 3))
                Else
                    Call SendData(ToIndex, UserIndex, 0, "VERSO" & MiRespuesta)
                    MiRespuesta = MiRespuesta & "0k1"
                    Call WriteVar(CharPath & UCase$(UserList(UserIndex).Name) & ".CHR", "STATS", "Respuesta", MiRespuesta)
                End If
            Else
            MiRespuesta = GetVar(CharPath & UCase$(UserList(UserIndex).Name) & ".CHR", "STATS", "Soporte")
                
                If Len(MiRespuesta) Then
                    Call SendData(ToIndex, UserIndex, 0, "||No respondida aún" & FONTTYPE_INFO)
                Else
                    Call SendData(ToIndex, UserIndex, 0, "||No has mandado ningun soporte!" & FONTTYPE_INFO)
                End If
            
            End If
            
        Exit Sub
    End If
    
    If UCase$(Left$(rdata, 8)) = "/SOPORTE" Then
    Call SendData(ToIndex, UserIndex, 0, "SHWSUP")
    End If
    
    
     'FuriusAO Sistema de soporte basico!
     If UCase$(Left$(rdata, 9)) = "/ZOPORTE " Then
        If SoporteDesactivado Then
            Call SendData(ToIndex, UserIndex, 0, "||El soporte se encuentra deshabilitado." & FONTTYPE_INFO)
            Exit Sub
        End If
        If Len(rdata) > 310 Then Exit Sub
        If InStr(rdata, "°") Then Exit Sub
        If InStr(rdata, "~") Then Exit Sub
       'If UserList(userindex).flags.Silenciado > 0 Then Exit Sub
        rdata = Right$(rdata, Len(rdata) - 9)
        'Call SendData(ToAdmins, 0, 0, "||" & UserList(userindex).Name & " >" & "" & "SOPORTE:" & rdata & FONTTYPE_FIGHT)
        'Call SendData(ToIndex, userindex, 0, "||El soporte fue enviado. Rogamos que tengas paciencia y aguardes a ser atendido por un GM. No escribas más de un mensaje sobre el mismo tema." & FONTTYPE_furius)
                
        Dim SoporteA As String
        
        SoporteA = GetVar(CharPath & UCase$(UserList(UserIndex).Name) & ".CHR", "STATS", "Respuesta")
        
        'SI HAY RESPUESTA Y NO ESTA LEIDA LE AVISA.
        If Len(SoporteA) > 0 And Right$(SoporteA, 3) <> "0k1" Then
        Call SendData(ToIndex, UserIndex, 0, "||Primero debes leer la respuesta de tu anterior soporte." & FONTTYPE_INFO)
        Exit Sub
        End If
        '/
        
        SoporteA = GetVar(CharPath & UCase$(UserList(UserIndex).Name) & ".CHR", "STATS", "Soporte")
        
        'SI MANDO SOPORTE ANTES Y TODAVIA NO LE RESPONDIERON TIENE QE ESPERAR
        If Len(SoporteA) > 0 And Right$(SoporteA, 3) <> "0k1" Then
        Call SendData(ToIndex, UserIndex, 0, "||Ya has mandado un soporte. Debes esperar la respuesta para enviar otro. " & FONTTYPE_INFO)
        Exit Sub
        End If
        '0K
        
        SoporteA = "Dia:" & Day(Now) & " Hora:" & Time & " - Soporte: " & Replace(Replace(rdata, ";", ":"), Chr$(13) & Chr$(10), Chr(32))
        
        
        
        
        Call WriteVar(CharPath & UCase$(UserList(UserIndex).Name) & ".CHR", "STATS", "Soporte", SoporteA)
        Call WriteVar(CharPath & UCase$(UserList(UserIndex).Name) & ".CHR", "STATS", "Respuesta", "")
        Soportes.Add (UserList(UserIndex).Name)
        Call SendData(ToIndex, UserIndex, 0, "||El soporte ha sido enviado con éxito. Gracias por utilizar nuestro sistema. Aguarde su respuesta." & FONTTYPE_INFO)
        Exit Sub
        End If
'FuriusAO SISTEMA DE SOPORTE BASICO.


If UCase$(Left$(rdata, 7)) = "/ENTRAR" Then
' aca le ponen las condiciones a su gusto, puede ser que sea mayor a tal lvl para entrar, que no puedan entrar invis, ni muertos, etc. a su gusto.
        If UserList(UserIndex).flags.Muerto Then
           Call SendData(ToIndex, UserIndex, 0, "||Estas muerto!, no puedes participar en el torneo." & FONTTYPE_INFO)
           Exit Sub
        End If
Call Torneos_Entra(UserIndex)
Exit Sub
End If

If UCase$(Left$(rdata, 12)) = "/MERCENARIO " Then
    rdata = Right$(rdata, Len(rdata) - 12)
    If Not ModoQuest Then Exit Sub
    If UserList(UserIndex).flags.Privilegios > 0 Then Exit Sub
    Select Case UCase$(rdata)
        Case "ALIANZA"
            tInt = 1
        Case "LORD THEK"
            tInt = 2
        Case Else
            Call SendData(ToIndex, UserIndex, 0, "||La estructura del comando es /MERCENARIO ALIANZA o /MERCENARIO LORD THEK." & FONTTYPE_FENIX)
            Exit Sub
    End Select
    
    Select Case UserList(UserIndex).Faccion.BandoOriginal
        Case Neutral
            If UserList(UserIndex).Faccion.Bando <> Neutral Then
                Call SendData(ToIndex, UserIndex, 0, "||Ya eres mercenario para " & ListaBandos(UserList(UserIndex).Faccion.Bando) & "." & FONTTYPE_FENIX)
                Exit Sub
            End If
        
        Case Else
            Select Case UserList(UserIndex).Faccion.Bando
                Case Neutral
                    If tInt = UserList(UserIndex).Faccion.BandoOriginal Then
                        Call SendData(ToIndex, UserIndex, 0, "||" & ListaBandos(tInt) & " no acepta desertores entre sus filas." & FONTTYPE_FENIX)
                        Exit Sub
                    End If
            
                Case UserList(UserIndex).Faccion.BandoOriginal
                    Call SendData(ToIndex, UserIndex, 0, "||Ya perteneces a " & ListaBandos(UserList(UserIndex).Faccion.Bando) & ", no puedes ofrecerte como mercenario." & FONTTYPE_FENIX)
                    Exit Sub
        
                Case Else
                    Call SendData(ToIndex, UserIndex, 0, "||Ya eres mercenario para " & ListaBandos(UserList(UserIndex).Faccion.Bando) & "." & FONTTYPE_FENIX)
                    Exit Sub
            End Select
    End Select
    Call SendData(ToIndex, UserIndex, 0, "||¡" & ListaBandos(tInt) & " te ha aceptado como un mercenario entre sus filas!" & FONTTYPE_FENIX)
    UserList(UserIndex).Faccion.Bando = tInt
    Call UpdateUserChar(UserIndex)
    Exit Sub
End If

If UserList(UserIndex).flags.Quest Then
    If UCase$(Left$(rdata, 3)) = "/M " Then
        rdata = Right$(rdata, Len(rdata) - 3)
        If Len(rdata) = 0 Then Exit Sub
        Select Case UserList(UserIndex).Faccion.Bando
            Case Real
                tStr = FONTTYPE_ARMADA
            Case Caos
                tStr = FONTTYPE_CAOS
        End Select
        Call SendData(ToAll, 0, 0, "||" & rdata & tStr)
        Exit Sub
    ElseIf UCase$(rdata) = "/TELEPLOC" Then
        Call WarpUserChar(UserIndex, UserList(UserIndex).flags.TargetMap, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY, True)
        Exit Sub
    ElseIf UCase$(rdata) = "/TRAMPA" Then
        Call ActivarTrampa(UserIndex)
        Exit Sub
    End If
End If

If UserList(UserIndex).flags.PuedeDenunciar Or UserList(UserIndex).flags.Privilegios > 0 Then
    If UCase$(Left$(rdata, 11)) = "/DENUNCIAS " Then
        rdata = Right$(rdata, Len(rdata) - 11)
        TIndex = NameIndex(rdata)
        
        If TIndex > 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||Denuncias por cheat: " & UserList(TIndex).flags.Denuncias & "." & FONTTYPE_INFO)
            Call SendData(ToIndex, UserIndex, 0, "||Denuncias por insultos: " & UserList(TIndex).flags.DenunciasInsultos & "." & FONTTYPE_INFO)
        Else
            Call SendData(ToIndex, UserIndex, 0, "1A")
        End If
        Exit Sub
    End If

    If UCase$(Left$(rdata, 6)) = "/DENC " Then
        rdata = Right$(rdata, Len(rdata) - 6)
        TIndex = NameIndex(rdata)
        
        If TIndex > 0 Then
            UserList(TIndex).flags.Denuncias = UserList(TIndex).flags.Denuncias + 1
            Call SendData(ToIndex, UserIndex, 0, "||Sumaste una denuncia por cheat a " & UserList(TIndex).Name & ". El usuario tiene acumuladas " & UserList(TIndex).flags.Denuncias & " denuncias." & FONTTYPE_INFO)
            Call LogGM(UserList(UserIndex).Name, "Sumo una denuncia por cheat a " & UserList(TIndex).Name & ".", UserList(UserIndex).flags.Privilegios = 1)
        Else
            If Not ExistePersonaje(rdata) Then
                Call SendData(ToIndex, UserIndex, 0, "||El personaje está offline y no se encuentra en la base de datos." & FONTTYPE_INFO)
                Exit Sub
            End If
            Call LogGM(UserList(UserIndex).Name, "Sumo una denuncia por cheat a " & rdata & ".", UserList(UserIndex).flags.Privilegios = 1)
            Call SendData(ToIndex, UserIndex, 0, "||Sumaste una denuncia por cheat a " & rdata & ". El usuario tiene acumuladas " & SumarDenuncia(rdata, 1) & " denuncias." & FONTTYPE_INFO)
        End If
        Exit Sub
    End If

    If UCase$(Left$(rdata, 6)) = "/DENI " Then
        rdata = Right$(rdata, Len(rdata) - 6)
        TIndex = NameIndex(rdata)
        
        If TIndex > 0 Then
            UserList(TIndex).flags.DenunciasInsultos = UserList(TIndex).flags.DenunciasInsultos + 1
            Call SendData(ToIndex, UserIndex, 0, "||Sumaste una denuncia por insultos a " & UserList(TIndex).Name & ". El usuario tiene acumuladas " & UserList(TIndex).flags.DenunciasInsultos & " denuncias." & FONTTYPE_INFO)
            Call LogGM(UserList(UserIndex).Name, "Sumo una denuncia por insultos a " & UserList(TIndex).Name & ".", UserList(UserIndex).flags.Privilegios = 1)
        Else
            If Not ExistePersonaje(rdata) Then
                Call SendData(ToIndex, UserIndex, 0, "||El personaje está offline y no se encuentra en la base de datos." & FONTTYPE_INFO)
                Exit Sub
            End If
            Call LogGM(UserList(UserIndex).Name, "Sumo una denuncia por insultos a " & rdata & ".", UserList(UserIndex).flags.Privilegios = 1)
            Call SendData(ToIndex, UserIndex, 0, "||Sumaste una denuncia por insultos a " & rdata & ". El usuario tiene acumuladas " & SumarDenuncia(rdata, 2) & " denuncias." & FONTTYPE_INFO)
        End If
        Exit Sub
    End If
End If

If UserList(UserIndex).flags.Privilegios = 0 Then Exit Sub

If UCase$(Left$(rdata, 8)) = "/PTORNEO" Then
Call SendData(ToIndex, UserIndex, 0, "PTOR")
Exit Sub
End If

If UCase$(rdata) = "/DAMESOS" Then
Dim LstU As String
    
    If Soportes.Count = 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||No hay soportes para ver." & FONTTYPE_INFO)
        Exit Sub
    End If

    For i = 1 To Soportes.Count
        LstU = LstU & "@" & Soportes.Item(i)
        Debug.Print Soportes.Item(i)
        DoEvents
    Next i

    LstU = Soportes.Count & LstU

    LstU = "SHWSOP@" & LstU
    Call SendData(ToIndex, UserIndex, 0, LstU)
    
End If

If UCase$(Left$(rdata, 7)) = "/BORSO " Then
rdata = Right$(rdata, Len(rdata) - 7)
Call WriteVar(CharPath & UCase$(rdata) & ".chr", "STATS ", "Soporte", "")
Call WriteVar(CharPath & UCase$(rdata) & ".chr", "STATS ", "Respuesta", "")
For i = 1 To Soportes.Count
If UCase$(Soportes.Item(i)) = UCase$(rdata) Then
    Soportes.Remove (i)
    Exit For
End If
DoEvents
Next i
Call SendData(ToIndex, UserIndex, 0, "||Soporte y respuesta borrados con éxito" & FONTTYPE_INFO)
Exit Sub
End If


If UCase$(Left$(rdata, 7)) = "/SOSDE " Then
rdata = Right$(rdata, Len(rdata) - 7)

Dim SosDe As String
SosDe = GetVar(CharPath & UCase$(rdata) & ".chr", "STATS", "Soporte")


    If Len(SosDe) > 0 Then
        Call SendData(ToIndex, UserIndex, 0, "SOPODE" & SosDe)
    Else
        Call SendData(ToIndex, UserIndex, 0, "||Error. Soporte no encontrado" & FONTTYPE_INFO)
    End If


End If

If UCase$(Left$(rdata, 7)) = "/RESOS " Then
rdata = Right$(rdata, Len(rdata) - 7)
Dim Persona, Respuesta As String
Persona = ReadField$(1, rdata, Asc(";")) 'GetVar(CharPath & UCase$(rdata) & ".chr", "STATS", "Soporte")
Respuesta = Replace(ReadField$(2, rdata, Asc(";")), Chr$(13) & Chr$(10), Chr(32))
If Len(Persona) = 0 Or Len(Respuesta) = 0 Then
    Call SendData(ToIndex, UserIndex, 0, "||Error en la respuesta" & FONTTYPE_INFO)
    Exit Sub
End If

Call WriteVar(CharPath & UCase$(Persona) & ".chr", "STATS", "Respuesta", Respuesta)
Call WriteVar(CharPath & UCase$(Persona) & ".chr", "STATS", "Soporte", GetVar(CharPath & UCase$(Persona) & ".chr", "STATS", "Soporte") & "0k1")


TIndex = NameIndex(Persona)
If TIndex > 0 Then
    Call SendData(ToIndex, TIndex, 0, "||Tu soporte ha sido respondido." & FONTTYPE_INFO)
    Call SendData(ToIndex, TIndex, 0, "TENSO")
End If
    
Call SendData(ToIndex, UserIndex, 0, "||Soporte respondido con éxito" & FONTTYPE_INFO)
    For i = 1 To Soportes.Count
    Debug.Print Soportes.Item(1)
    
        If UCase$(Soportes.Item(i)) = UCase$(Persona) Then
            Soportes.Remove (i)
            Exit For
        End If
        DoEvents
    Next i


End If

If UCase$(Left$(rdata, 4)) = "/GO " Then
    rdata = Right$(rdata, Len(rdata) - 4)
    mapa = val(ReadField(1, rdata, 32))
    If Not MapaValido(mapa) Then Exit Sub
    If UserList(UserIndex).flags.Privilegios = 1 And MapInfo(mapa).Pk Then Exit Sub
    Call WarpUserChar(UserIndex, mapa, 50, 50, True)
    Call SendData(ToIndex, UserIndex, 0, "2B" & UserList(UserIndex).Name)
    Call LogGM(UserList(UserIndex).Name, "Transporto a " & UserList(UserIndex).Name & " hacia " & "Mapa" & mapa & " X:" & X & " Y:" & Y, (UserList(UserIndex).flags.Privilegios = 1))
    Exit Sub
End If

If UCase$(Left$(rdata, 10)) = "/DARHONOR " Then
    rdata = Right$(rdata, Len(rdata) - 10)
    TIndex = UserList(UserIndex).flags.TargetUser
    If TIndex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||Debes seleccionar al Jugador para Darle sus Puntos!" & FONTTYPE_INFO)
        Exit Sub
    End If
    If rdata < 0 Then
    Call SendData(ToIndex, UserIndex, 0, "||No podes dar puntos negativos." & FONTTYPE_FENIX)
    Exit Sub
    End If
    If rdata >= 10000 Then
        Call SendData(ToIndex, UserIndex, 0, "||No puedes Entregar mas de 5 Puntos de Reto" & FONTTYPE_INFO)
        Exit Sub
    End If
    Call SendData(ToAll, 0, 0, "||" & UserList(UserList(UserIndex).flags.TargetUser).Name & " gano " & rdata & " puntos de Reto" & FONTTYPE_FENIX)
    UserList(UserList(UserIndex).flags.TargetUser).flags.Reto = UserList(UserList(UserIndex).flags.TargetUser).flags.Reto + rdata
    Call LogGM(UserList(UserIndex).Name, "Puntos de Reto: " & rdata & UserList(TIndex).Name & " Map:" & UserList(UserIndex).POS.Map & " X:" & UserList(UserIndex).POS.X & " Y:" & UserList(UserIndex).POS.Y, False)
    Exit Sub
End If

If UCase$(Left$(rdata, 10)) = "/DARPUNTO " Then
    rdata = Right$(rdata, Len(rdata) - 10)
    TIndex = UserList(UserIndex).flags.TargetUser
    If TIndex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||Debes seleccionar al Jugador para Darle sus Puntos!" & FONTTYPE_INFO)
        Exit Sub
    End If
    If rdata < 0 Then
    Call SendData(ToIndex, UserIndex, 0, "||No podes dar puntos negativos." & FONTTYPE_FENIX)
    Exit Sub
    End If
    Call SendData(ToAll, 0, 0, "||" & UserList(UserList(UserIndex).flags.TargetUser).Name & " gano " & rdata & " puntos de Canje" & FONTTYPE_FENIX)
    UserList(UserList(UserIndex).flags.TargetUser).flags.Canje = UserList(UserList(UserIndex).flags.TargetUser).flags.Canje + rdata
    Call LogGM(UserList(UserIndex).Name, "Puntos de Canje: " & rdata & UserList(TIndex).Name & " Map:" & UserList(UserIndex).POS.Map & " X:" & UserList(UserIndex).POS.X & " Y:" & UserList(UserIndex).POS.Y, False)
    Exit Sub
End If
 
If UCase$(Left$(rdata, 12)) = "/SACARPUNTO " Then
    rdata = Right$(rdata, Len(rdata) - 12)
    TIndex = UserList(UserIndex).flags.TargetUser
    If rdata < 0 Then
    Call SendData(ToIndex, UserIndex, 0, "||No podes sacar cantidades negativas" & FONTTYPE_FENIX)
    Exit Sub
    End If
    If TIndex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||Debes seleccionar al Jugador para Sacarle sus Puntos!" & FONTTYPE_INFO)
        Exit Sub
    End If
    If UserList(TIndex).flags.Canje < rdata Then
        Call SendData(ToIndex, UserIndex, 0, "||No puedes Sacar esa Cantidad de Puntos, Genera Variable Muerta!" & FONTTYPE_INFO)
        Exit Sub
    End If
    Call SendData(ToIndex, UserIndex, 0, "||Le has quitado " & rdata & " puntos de canje a " & UserList(TIndex).Name & "!" & FONTTYPE_FENIX)
    UserList(UserList(UserIndex).flags.TargetUser).flags.Canje = UserList(UserList(UserIndex).flags.TargetUser).flags.Canje - rdata
    Call SendData(ToIndex, TIndex, 0, "||" & UserList(UserIndex).Name & " te ha quitado " & rdata & " puntos de canje!" & FONTTYPE_INFO)
    Call LogGM(UserList(UserIndex).Name, "Restó puntos de canje: " & rdata & UserList(TIndex).Name & " Map:" & UserList(UserIndex).POS.Map & " X:" & UserList(UserIndex).POS.X & " Y:" & UserList(UserIndex).POS.Y, False)
    Exit Sub
End If


If UCase$(Left$(rdata, 5)) = "/SUM " Then
    rdata = Right$(rdata, Len(rdata) - 5)
    
    TIndex = NameIndex(rdata)
    If TIndex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "1A")
        Exit Sub
    End If
    
    If UserList(UserIndex).flags.Privilegios < UserList(TIndex).flags.Privilegios And UserList(TIndex).flags.AdminInvisible = 1 Then
        Call SendData(ToIndex, UserIndex, 0, "1A")
        Exit Sub
    End If
    
    If UserList(UserIndex).flags.Privilegios = 1 And UserList(TIndex).POS.Map <> UserList(UserIndex).POS.Map Then Exit Sub
    
    Call SendData(ToIndex, UserIndex, 0, "%Z" & UserList(TIndex).Name)
    Call WarpUserChar(TIndex, UserList(UserIndex).POS.Map, UserList(UserIndex).POS.X, UserList(UserIndex).POS.Y + 1, True)
    
    Call LogGM(UserList(UserIndex).Name, "/SUM " & UserList(TIndex).Name & " Map:" & UserList(UserIndex).POS.Map & " X:" & UserList(UserIndex).POS.X & " Y:" & UserList(UserIndex).POS.Y, False)
    Exit Sub
End If

If UCase$(Left$(rdata, 5)) = "/IRA " Then
    rdata = Right$(rdata, Len(rdata) - 5)
    
    TIndex = NameIndex(rdata)
    If TIndex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "1A")
        Exit Sub
    End If
    
    If ((UserList(TIndex).flags.Privilegios > UserList(UserIndex).flags.Privilegios And UserList(TIndex).flags.AdminInvisible = 1)) Then
        Call SendData(ToIndex, UserIndex, 0, "1A")
        Exit Sub
    End If

    If UserList(TIndex).flags.AdminInvisible And Not UserList(UserIndex).flags.AdminInvisible Then Call DoAdminInvisible(UserIndex)

    Call WarpUserChar(UserIndex, UserList(TIndex).POS.Map, UserList(TIndex).POS.X + 1, UserList(TIndex).POS.Y + 1, True)
    
    Call LogGM(UserList(UserIndex).Name, "/IRA " & UserList(TIndex).Name & " Mapa:" & UserList(TIndex).POS.Map & " X:" & UserList(TIndex).POS.X & " Y:" & UserList(TIndex).POS.Y, (UserList(UserIndex).flags.Privilegios = 1))
    Exit Sub
End If

If UCase$(rdata) = "/TRABAJANDO" Then
    For LoopC = 1 To LastUser
        If Len(UserList(LoopC).Name) > 0 And UserList(LoopC).flags.Trabajando Then
            DummyInt = DummyInt + 1
            tStr = tStr & UserList(LoopC).Name & ", "
        End If
    Next
    If Len(tStr) > 0 Then
        tStr = Left$(tStr, Len(tStr) - 2)
        Call SendData(ToIndex, UserIndex, 0, "||Usuarios trabajando: " & tStr & FONTTYPE_INFO)
        Call SendData(ToIndex, UserIndex, 0, "||Número de usuarios trabajando: " & DummyInt & "." & FONTTYPE_INFO)
    Else
        Call SendData(ToIndex, UserIndex, 0, "%)")
    End If
    Exit Sub
End If
If UCase$(Left$(rdata, 8)) = "/CARCEL " Then
    
    rdata = Right$(rdata, Len(rdata) - 8)
    
    Name = ReadField(1, rdata, 32)
    i = val(ReadField(1, rdata, 32))
    Name = Right$(rdata, Len(rdata) - (Len(Name) + 1))
    
    TIndex = NameIndex(Name)
    
    If TIndex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "1A")
        Exit Sub
    End If
    
    If UserList(TIndex).flags.Privilegios > UserList(UserIndex).flags.Privilegios Then
        Call SendData(ToIndex, UserIndex, 0, "1B")
        Exit Sub
    End If
    
    If i > 120 Then
        Call SendData(ToIndex, UserIndex, 0, "1C")
        Exit Sub
    End If
    
    Call Encarcelar(TIndex, i, UserList(UserIndex).Name)
    
    Exit Sub
End If


If UserList(UserIndex).flags.Privilegios < 2 Then Exit Sub

If UCase$(Left$(rdata, 4)) = "/REM" Then
    rdata = Right$(rdata, Len(rdata) - 5)
    Call LogGM(UserList(UserIndex).Name, "Comentario: " & rdata, (UserList(UserIndex).flags.Privilegios = 1))
    Call SendData(ToIndex, UserIndex, 0, "||Comentario salvado..." & FONTTYPE_INFO)
    Exit Sub
End If

If UCase$(Left$(rdata, 7)) = "/STAFF " Then
    rdata = Right$(rdata, Len(rdata) - 7)
    Call LogGM(UserList(UserIndex).Name, "Mensaje a Gms:" & rdata, (UserList(UserIndex).flags.Privilegios = 1))
    If Len(rdata) > 0 Then
        Call SendData(ToAdmins, 0, 0, "||" & UserList(UserIndex).Name & "> " & rdata & "~255~255~255~0~1")
    End If
    Exit Sub
End If


If UCase$(Left$(rdata, 5)) = "/HORA" Then
    Call LogGM(UserList(UserIndex).Name, "Hora.", (UserList(UserIndex).flags.Privilegios = 1))
    rdata = Right$(rdata, Len(rdata) - 5)
    Call SendData(ToAll, 0, 0, "||Hora: " & Time & " " & Date & FONTTYPE_INFO)
    Exit Sub
End If

If UCase$(rdata) = "/ONLINEGM" Then
        For LoopC = 1 To LastUser
            If Len(UserList(LoopC).Name) > 0 Then
                If UserList(LoopC).flags.Privilegios > 0 And (UserList(LoopC).flags.Privilegios <= UserList(UserIndex).flags.Privilegios Or UserList(LoopC).flags.AdminInvisible = 0) Then
                    tStr = tStr & UserList(LoopC).Name & ", "
                End If
            End If
        Next
        If Len(tStr) > 0 Then
            tStr = Left$(tStr, Len(tStr) - 2)
            Call SendData(ToIndex, UserIndex, 0, "||" & tStr & FONTTYPE_INFO)
        Else
            Call SendData(ToIndex, UserIndex, 0, "%P")
        End If
        Exit Sub
End If

If UCase$(Left$(rdata, 7)) = "/DONDE " Then
    rdata = Right$(rdata, Len(rdata) - 7)
    TIndex = NameIndex(rdata)
    If TIndex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "1A")
        Exit Sub
    End If
    
    If UserList(TIndex).flags.Privilegios > UserList(UserIndex).flags.Privilegios And UserList(TIndex).flags.AdminInvisible = 1 Then
        Call SendData(ToIndex, UserIndex, 0, "1A")
        Exit Sub
    End If
    
    Call SendData(ToIndex, UserIndex, 0, "||Ubicacion de " & UserList(TIndex).Name & ": " & UserList(TIndex).POS.Map & ", " & UserList(TIndex).POS.X & ", " & UserList(TIndex).POS.Y & "." & FONTTYPE_INFO)
    Call LogGM(UserList(UserIndex).Name, "/Donde", (UserList(UserIndex).flags.Privilegios = 1))
    Exit Sub
End If

If UCase$(Left$(rdata, 6)) = "/NENE " Then
    rdata = Right$(rdata, Len(rdata) - 6)
    If MapaValido(val(rdata)) Then
        Call SendData(ToIndex, UserIndex, 0, "NENE" & NPCHostiles(val(rdata)))
        Call LogGM(UserList(UserIndex).Name, "Numero enemigos en mapa " & rdata, (UserList(UserIndex).flags.Privilegios = 1))
    End If
    Exit Sub
End If

If UCase$(rdata) = "/VENTAS" Then
    Call SendData(ToIndex, UserIndex, 0, "/X" & DineroTotalVentas & "," & NumeroVentas)
    Exit Sub
End If

If UCase$(rdata) = "/TELEPLOC" Then
    Call WarpUserChar(UserIndex, UserList(UserIndex).flags.TargetMap, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY, True)
    Call LogGM(UserList(UserIndex).Name, "/TELEPLOC a x:" & UserList(UserIndex).flags.TargetX & " Y:" & UserList(UserIndex).flags.TargetY & " Map:" & UserList(UserIndex).POS.Map, (UserList(UserIndex).flags.Privilegios = 1))
    Exit Sub
End If

If UCase$(rdata) = "/DESCONGELAR" Then
    Call Congela(True)
    Exit Sub
End If

If UCase$(Left$(rdata, 9)) = "/VIGILAR " Then
    rdata = Right$(rdata, Len(rdata) - 9)
    TIndex = NameIndex(rdata)
    If TIndex > 0 Then
        If TIndex = UserIndex Then
            Call SendData(ToIndex, UserIndex, 0, "||No puedes vigilarte a ti mismo." & FONTTYPE_INFO)
            Exit Sub
        End If
        If UserList(TIndex).flags.Privilegios >= UserList(UserIndex).flags.Privilegios Then
            Call SendData(ToIndex, UserIndex, 0, "||No puedes vigilar a alguien con igual o mayor jerarquia que tú." & FONTTYPE_INFO)
            Exit Sub
        End If
        If YaVigila(TIndex, UserIndex) Then
            Call SendData(ToIndex, UserIndex, 0, "||Dejaste de vigilar a " & UserList(TIndex).Name & "." & FONTTYPE_INFO)
            If Not EsVigilado(TIndex) Then Call SendData(ToIndex, TIndex, 0, "VIG")
            Exit Sub
        End If
        If Not EsVigilado(TIndex) Then Call SendData(ToIndex, TIndex, 0, "VIG")
        Call SendData(ToIndex, UserIndex, 0, "||Estás vigilando a " & UserList(TIndex).Name & "." & FONTTYPE_INFO)
        For i = 1 To 10
            If UserList(TIndex).flags.Espiado(i) = 0 Then
                UserList(TIndex).flags.Espiado(i) = UserIndex
                Exit For
            End If
        Next
        If i = 11 Then
            Call SendData(ToIndex, UserIndex, 0, "||Demasiados GM's están vigilando a este usuario." & FONTTYPE_INFO)
            Exit Sub
        End If
    Else
        Call SendData(ToIndex, UserIndex, 0, "1A")
    End If
    Exit Sub
End If

If UCase$(Left$(rdata, 7)) = "/VERPC " Then
    rdata = Right$(rdata, Len(rdata) - 7)
    TIndex = NameIndex(rdata)
    If TIndex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "1A")
        Exit Sub
    End If
    If UserList(TIndex).flags.Privilegios > UserList(UserIndex).flags.Privilegios And UserList(UserIndex).flags.AdminInvisible = 1 Then
        Call SendData(ToIndex, UserIndex, 0, "1A")
        Exit Sub
    End If
    If UserList(TIndex).flags.Privilegios >= UserList(UserIndex).flags.Privilegios Then
        Call SendData(ToIndex, UserIndex, 0, "||No puedes ver la PC de un GM con mayor jerarquia." & FONTTYPE_FIGHT)
        Exit Sub
    End If
    UserList(TIndex).flags.EsperandoLista = UserIndex
    Call SendData(ToIndex, TIndex, 0, "VPRC")
End If

If UCase$(Left$(rdata, 7)) = "/TELEP " Then
    rdata = Right$(rdata, Len(rdata) - 7)
    mapa = val(ReadField(2, rdata, 32))
    If Not MapaValido(mapa) Then Exit Sub
    Name = ReadField(1, rdata, 32)
    If Len(Name) = 0 Then Exit Sub
    If UCase$(Name) <> "YO" Then
        If UserList(UserIndex).flags.Privilegios = 1 Then
            Exit Sub
        End If
        TIndex = NameIndex(Name)
    Else
        TIndex = UserIndex
    End If
    X = val(ReadField(3, rdata, 32))
    Y = val(ReadField(4, rdata, 32))
    If Not InMapBounds(X, Y) Then Exit Sub
    If TIndex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "1A")
        Exit Sub
    End If
    If UserList(TIndex).flags.Privilegios > UserList(UserIndex).flags.Privilegios And UserList(UserIndex).flags.AdminInvisible = 1 Then
        Call SendData(ToIndex, UserIndex, 0, "1A")
        Exit Sub
    End If
    Call WarpUserChar(TIndex, mapa, X, Y, True)
    Call SendData(ToIndex, TIndex, 0, "||" & UserList(UserIndex).Name & " te ha transportado." & FONTTYPE_INFO)
    Call LogGM(UserList(UserIndex).Name, "Transporto a " & UserList(TIndex).Name & " hacia " & "Mapa" & mapa & " X:" & X & " Y:" & Y, (UserList(UserIndex).flags.Privilegios = 1))
    Exit Sub
End If


If UCase$(Left$(rdata, 4)) = "/GO " Then
    rdata = Right$(rdata, Len(rdata) - 4)
    mapa = val(ReadField(1, rdata, 32))
    If Not MapaValido(mapa) Then Exit Sub
    Call WarpUserChar(UserIndex, mapa, 50, 50, True)
    Call SendData(ToIndex, UserIndex, 0, "2B" & UserList(UserIndex).Name)
    Call LogGM(UserList(UserIndex).Name, "Transporto a " & UserList(UserIndex).Name & " hacia " & "Mapa" & mapa & " X:" & X & " Y:" & Y, (UserList(UserIndex).flags.Privilegios = 1))
    Exit Sub
End If

If UCase$(rdata) = "/OMAP" Then
    For LoopC = 1 To MapInfo(UserList(UserIndex).POS.Map).NumUsers
        If UserList(MapInfo(UserList(UserIndex).POS.Map).UserIndex(LoopC)).flags.Privilegios <= UserList(UserIndex).flags.Privilegios Then
            tStr = tStr & UserList(MapInfo(UserList(UserIndex).POS.Map).UserIndex(LoopC)).Name & ","
        End If
    Next
    If Len(tStr) > 0 Then
        tStr = Left$(tStr, Len(tStr) - 1)
        Call SendData(ToIndex, UserIndex, 0, "||Usuarios en este mapa: " & tStr & "." & FONTTYPE_INFO)
    Else
        Call SendData(ToIndex, UserIndex, 0, "%R")
    End If
    Exit Sub
End If

If UCase$(rdata) = "/PANELGM" Then
    Call SendData(ToIndex, UserIndex, 0, "PGM" & UserList(UserIndex).flags.Privilegios)
    Exit Sub
End If

If UCase$(rdata) = "/CMAP" Then
    If MapInfo(UserList(UserIndex).POS.Map).NumUsers Then
        Call SendData(ToIndex, UserIndex, 0, "||Hay " & MapInfo(UserList(UserIndex).POS.Map).NumUsers & " usuarios en este mapa." & FONTTYPE_INFO)
    Else
        Call SendData(ToIndex, UserIndex, 0, "%R")
    End If

    Exit Sub
End If

If UCase$(Left$(rdata, 7)) = "/TORNEO" Then
    If entorneo = 0 Then
        entorneo = 1
        If FileExist(App.Path & "/logs/torneo.log", vbNormal) Then Kill (App.Path & "/logs/torneo.log")
        Call SendData(ToIndex, UserIndex, 0, "||Has activado el torneo" & FONTTYPE_GANOTOQUE)
        Call SendData(ToAll, 0, 0, "||Han abierto las inscripciones para participar en el torneo, escribe /PARTICIPAR para ingresar." & FONTTYPE_FENIX)
    Else
        entorneo = 0
        Call SendData(ToIndex, UserIndex, 0, "||Has desactivado el torneo" & FONTTYPE_GANOTOQUE)
        Call SendData(ToAll, 0, 0, "||Inscripciones cerradas." & FONTTYPE_FENIX)
    End If
    Exit Sub
End If

If UCase$(Left$(rdata, 10)) = "/VERTORNEO" Then
    Dim stri As String
    Dim Jugadores As Integer
    Dim jugador As Integer
    stri = ""
    Jugadores = val(GetVar(App.Path & "/logs/torneo.log", "CANTIDAD", "CANTIDAD"))
    For jugador = 1 To Jugadores
        stri = stri & GetVar(App.Path & "/logs/torneo.log", "JUGADORES", "JUGADOR" & jugador) & "@"
    Next
    'Call SendData(ToIndex, UserIndex, 0, "||Quieren participar: " & stri & FONTTYPE_ROSITA)
    Call SendData(ToIndex, UserIndex, 0, "PPT" & " @" & stri)
    Exit Sub
End If

If UCase$(rdata) = "/SOPORTEACTIVADO" Then
SoporteDesactivado = Not SoporteDesactivado
Call SendData(ToIndex, UserIndex, 0, "||El soporte está desactivado : " & SoporteDesactivado & FONTTYPE_INFO)
Exit Sub
End If

If UCase$(rdata) = "/INVISIBLE" Then
    Call DoAdminInvisible(UserIndex)
    Call LogGM(UserList(UserIndex).Name, "/INVISIBLE", (UserList(UserIndex).flags.Privilegios = 1))
    Exit Sub
End If


If UCase$(Left$(rdata, 6)) = "/INFO " Then
    Call LogGM(UserList(UserIndex).Name, rdata, False)
    
    rdata = Right$(rdata, Len(rdata) - 6)
    
    TIndex = NameIndex(rdata)
    
    If TIndex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "1A")
        Exit Sub
    End If

    SendUserSTAtsTxt UserIndex, TIndex
    Call SendData(ToIndex, UserIndex, 0, "||Mail: " & UserList(TIndex).Email & FONTTYPE_INFO)
    Call SendData(ToIndex, UserIndex, 0, "||Ip: " & UserList(TIndex).ip & FONTTYPE_INFO)

    Exit Sub
End If

If UCase$(Left$(rdata, 8)) = "/IPNICK " Then
    rdata = Right$(rdata, Len(rdata) - 8)






    tStr = ""
    For LoopC = 1 To LastUser
        If UserList(LoopC).ip = rdata And Len(UserList(LoopC).Name) > 0 And UserList(LoopC).flags.UserLogged Then
            If (UserList(UserIndex).flags.Privilegios > 0 And UserList(LoopC).flags.Privilegios = 0) Or (UserList(UserIndex).flags.Privilegios = 3) Then
                tStr = tStr & UserList(LoopC).Name & ", "
            End If
        End If
    Next
    Call SendData(ToIndex, UserIndex, 0, "||Los personajes con ip " & rdata & " son: " & tStr & FONTTYPE_INFO)
    Exit Sub
End If

If UCase$(Left$(rdata, 10)) = "/MAILNICK " Then
    rdata = Right$(rdata, Len(rdata) - 10)






    tStr = ""
    For LoopC = 1 To LastUser
        If UCase$(UserList(LoopC).ip) = UCase$(rdata) And Len(UserList(LoopC).Name) > 0 And UserList(LoopC).flags.UserLogged Then
            If (UserList(UserIndex).flags.Privilegios > 0 And UserList(LoopC).flags.Privilegios = 0) Or (UserList(UserIndex).flags.Privilegios = 3) Then
                tStr = tStr & UserList(LoopC).Name & ", "
            End If
        End If
    Next
    Call SendData(ToIndex, UserIndex, 0, "||Los personajes con mail " & rdata & " son: " & tStr & FONTTYPE_INFO)
    Exit Sub
End If


If UCase$(Left$(rdata, 5)) = "/INV " Then
    Call LogGM(UserList(UserIndex).Name, rdata, False)
    
    rdata = Right$(rdata, Len(rdata) - 5)
    
    TIndex = NameIndex(rdata)
    
    If TIndex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "1A")
        Exit Sub
    End If

    SendUserInvTxt UserIndex, TIndex
    Exit Sub
End If


If UCase$(Left$(rdata, 8)) = "/SKILLS " Then
    Call LogGM(UserList(UserIndex).Name, rdata, False)
    
    rdata = Right$(rdata, Len(rdata) - 8)
    
    TIndex = NameIndex(rdata)
    
    If TIndex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "1A")
        Exit Sub
    End If

    SendUserSkillsTxt UserIndex, TIndex
    Exit Sub
End If

If UCase$(Left$(rdata, 5)) = "/ATR " Then
    Call LogGM(UserList(UserIndex).Name, rdata, False)
    
    rdata = Right$(rdata, Len(rdata) - 5)
    
    TIndex = NameIndex(rdata)
    
    If TIndex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "1A")
        Exit Sub
    End If

    Call SendData(ToIndex, UserIndex, 0, "||Atributos de " & UserList(TIndex).Name & FONTTYPE_INFO)
    For i = 1 To NUMATRIBUTOS
        Call SendData(ToIndex, UserIndex, 0, "|| " & AtributosNames(i) & " = " & UserList(TIndex).Stats.UserAtributosBackUP(1) & FONTTYPE_INFO)
    Next
    Exit Sub
End If

If UCase$(Left$(rdata, 9)) = "/REVIVIR " Then
    rdata = Right$(rdata, Len(rdata) - 9)
    Name = rdata
    If UCase$(Name) <> "YO" Then
        TIndex = NameIndex(Name)
    Else
        TIndex = UserIndex
    End If
    If TIndex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "1A")
        Exit Sub
    End If
    Call RevivirUsuarioNPC(TIndex)
    Call SendData(ToIndex, TIndex, 0, "%T" & UserList(UserIndex).Name)
    Call LogGM(UserList(UserIndex).Name, "Resucito a " & UserList(TIndex).Name, False)
    Exit Sub
End If

If UCase$(Left$(rdata, 6)) = "/BANT " Then
    rdata = Right$(rdata, Len(rdata) - 6)

    Arg1 = ReadField(1, rdata, 64)
    Name = ReadField(2, rdata, 64)
    i = val(ReadField(3, rdata, 64))
    
    If Len(Arg1) = 0 Or Len(Name) = 0 Or i = 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||La estructura del comando es /BANT CAUSA@NICK@DIAS." & FONTTYPE_FENIX)
        Exit Sub
    End If
    
    TIndex = NameIndex(Name)
    
    If TIndex > 0 Then
        If UserList(TIndex).flags.Privilegios > UserList(UserIndex).flags.Privilegios Then
            Call SendData(ToIndex, UserIndex, 0, "1B")
            Exit Sub
        End If
        
        Call BanTemporal(Name, i, Arg1, UserList(UserIndex).Name)
        Call SendData(ToAdmins, 0, 0, "%X" & UserList(UserIndex).Name & "," & UserList(TIndex).Name)
        
        UserList(TIndex).flags.Ban = 1
        Call WarpUserChar(TIndex, ULLATHORPE.Map, ULLATHORPE.X, ULLATHORPE.Y)
        
        Call CloseSocket(TIndex)
    Else
        If Not FileExist(CharPath & UCase$(Name) & ".chr", vbNormal) = False Then
            Call SendData(ToIndex, UserIndex, 0, "||Offline, baneando" & FONTTYPE_INFO)
            
            If GetVar(CharPath & Name & ".chr", "FLAGS", "Ban") <> "0" Then
                Call SendData(ToIndex, UserIndex, 0, "||El personaje ya se encuentra baneado." & FONTTYPE_INFO)
                Exit Sub
            End If
            
            Call BanTemporal(Name, i, Arg1, UserList(UserIndex).Name)
            
            Call ChangeBan(Name, 1)
            Call ChangePos(Name)
            
            Call SendData(ToAdmins, 0, 0, "%X" & UserList(UserIndex).Name & "," & Name)
        Else
            Call SendData(ToIndex, UserIndex, 0, "||El usuario no existe." & FONTTYPE_INFO)
        End If
    End If

    Exit Sub
End If

If UCase$(Left$(rdata, 7)) = "/ECHAR " Then

    rdata = Right$(rdata, Len(rdata) - 7)
    TIndex = NameIndex(rdata)

    If TIndex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "1E")
        Exit Sub
    End If
    
    If TIndex = UserIndex Then Exit Sub
    
    If UserList(TIndex).flags.Privilegios > UserList(UserIndex).flags.Privilegios Then
        Call SendData(ToIndex, UserIndex, 0, "1F")
        Exit Sub
    End If
        
    Call SendData(ToAdmins, 0, 0, "%U" & UserList(UserIndex).Name & "," & UserList(TIndex).Name)
    Call LogGM(UserList(UserIndex).Name, "Echo a " & UserList(TIndex).Name, False)
    Call CloseSocket(TIndex)
    Exit Sub
End If

If UCase$(Left$(rdata, 5)) = "/BAN " Then
    Dim Razon As String
    rdata = Right$(rdata, Len(rdata) - 5)
    Razon = ReadField(1, rdata, Asc("@"))
    Name = ReadField(2, rdata, Asc("@"))
    TIndex = NameIndex(Name)
    '/ban motivo@nombre
    If TIndex Then
        If TIndex = UserIndex Then Exit Sub
        Name = UserList(TIndex).Name
        If UserList(TIndex).flags.Privilegios > UserList(UserIndex).flags.Privilegios Then
            Call SendData(ToIndex, UserIndex, 0, "%V")
            Exit Sub
        End If
        
        Call LogBan(TIndex, UserIndex, Razon)
        UserList(TIndex).flags.Ban = 1
        
        If UserList(TIndex).flags.Privilegios Then
            UserList(UserIndex).flags.Ban = 1
            Call SendData(ToAdmins, 0, 0, "%W" & UserList(UserIndex).Name)
            Call LogBan(UserIndex, UserIndex, "Baneado por banear a otro GM.")
            Call CloseSocket(UserIndex)
        End If
        
        Call SendData(ToAdmins, 0, 0, "%X" & UserList(UserIndex).Name & "," & UserList(TIndex).Name)
        Call SendData(ToAdmins, 0, 0, "||IP: " & UserList(TIndex).ip & " Mail: " & UserList(TIndex).Email & "." & FONTTYPE_FIGHT)

        Call CloseSocket(TIndex)
    Else
        If Not FileExist(CharPath & UCase$(Name) & ".chr", vbNormal) = False Then
            Call ChangeBan(Name, 1)
            Call LogBanOffline(UCase$(Name), UserIndex, Razon)
            Call SendData(ToAdmins, 0, 0, "%X" & UserList(UserIndex).Name & "," & Name)
        Else
            Call SendData(ToIndex, UserIndex, 0, "||El usuario no existe." & FONTTYPE_INFO)
        End If
        
    End If
    Exit Sub
End If


If UCase$(Left$(rdata, 7)) = "/UNBAN " Then
    rdata = Right$(rdata, Len(rdata) - 7)
    If Not FileExist(CharPath & UCase$(rdata) & ".chr", vbNormal) = False Then
        Call ChangeBan(rdata, 0)
        Call SendData(ToIndex, UserIndex, 0, "||" & rdata & " unbanned." & FONTTYPE_INFO)
        For i = 1 To Baneos.Count
            If Baneos(i).Name = UCase$(rdata) Then
                Call Baneos.Remove(i)
                Exit Sub
            End If
        Next
    Else
        Call SendData(ToIndex, UserIndex, 0, "||El usuario no existe" & FONTTYPE_INFO)
    End If
    Exit Sub
End If

'If UCase$(Left$(rdata, 7)) = "/UNBAN " Then
'    rdata = Right$(rdata, Len(rdata) - 7)
'
'    If Not ExistePersonaje(rdata) Then Exit Sub
'
'    Call ChangeBan(rdata, 0)
'
'    Call LogGM(UserList(UserIndex).Name, "/UNBAN a " & rdata, False)
'
'    Call SendData(ToIndex, UserIndex, 0, "%Y" & rdata)
'
'    For i = 1 To Baneos.Count
'        If Baneos(i).Name = UCase$(rdata) Then
'            Call Baneos.Remove(i)
'            Exit Sub
'        End If
'    Next
'
'    Exit Sub
'End If


If UCase$(rdata) = "/SEGUIR" Then
    If UserList(UserIndex).flags.TargetNpc Then
        Call DoFollow(UserList(UserIndex).flags.TargetNpc, UserIndex)
    End If
    Exit Sub
End If


If UCase$(Left$(rdata, 3)) = "/CC" Then
   Call EnviarSpawnList(UserIndex)
   Exit Sub
End If


If UCase$(Left$(rdata, 3)) = "SPA" Then
    rdata = Right$(rdata, Len(rdata) - 3)
    
    If val(rdata) > 0 And val(rdata) < UBound(SpawnList) + 1 Then _
          Call SpawnNpc(SpawnList(val(rdata)).NpcIndex, UserList(UserIndex).POS, True, False)
          
          Call LogGM(UserList(UserIndex).Name, "Sumoneo " & SpawnList(val(rdata)).NpcName, False)
          
    Exit Sub
End If

If UCase$(rdata) = "/RESETINV" Then
    rdata = Right$(rdata, Len(rdata) - 9)
    If UserList(UserIndex).flags.TargetNpc = 0 Then Call SendData(ToIndex, UserIndex, 0, "||No podes resetear" & FONTTYPE_TALK)
    Exit Sub
    Call ResetNpcInv(UserList(UserIndex).flags.TargetNpc)
    Call LogGM(UserList(UserIndex).Name, "/RESETINV " & Npclist(UserList(UserIndex).flags.TargetNpc).Name, False)
    Exit Sub
End If

If UCase$(rdata) = "/LIMPIAR" Then
    Call LimpiarMundo
    Exit Sub
End If

If UCase$(Left$(rdata, 6)) = "/RMSG " Then '[LmB/Mercury 2008]
rdata = Right$(rdata, Len(rdata) - 6)
If UserList(UserIndex).flags.Privilegios > 4 Then Exit Sub
Call LogGM(UserList(UserIndex).Name, "Mensaje Broadcast:" & rdata, False)
If rdata <> "" Then
Call SendData(ToAll, 0, 0, "|$" & UserList(UserIndex).Name & "> " & rdata)
End If
Exit Sub
End If '[LmB/Mercury 2008]

If UCase$(Left$(rdata, 7)) = "/RMSGT " Then
    rdata = Right$(rdata, Len(rdata) - 7)
    If UCase$(rdata) = "NO" Then
        Call SendData(ToAdmins, 0, 0, "||" & UserList(UserIndex).Name & " ha anulado la repetición del mensaje: " & MensajeRepeticion & "." & FONTTYPE_FENIX)
        IntervaloRepeticion = 0
        TiempoRepeticion = 0
        MensajeRepeticion = ""
        Exit Sub
    End If
    tName = ReadField(1, rdata, 64)
    tInt = ReadField(2, rdata, 64)
    Prueba1 = ReadField(3, rdata, 64)
    If Len(tName) = 0 Or val(Prueba1) = 0 Or (Prueba1 >= tInt And tInt <> 0) Then
        Call SendData(ToIndex, UserIndex, 0, "||La estructura del comando es: /RMSGT MENSAJE@TIEMPO TOTAL@INTERVALO DE REPETICION." & FONTTYPE_INFO)
        Exit Sub
    End If
    If val(tInt) > 10000 Or val(Prueba1) > 10000 Then
        Call SendData(ToIndex, UserIndex, 0, "||La cantidad de tiempo establecida es demasiado grande." & FONTTYPE_INFO)
        Exit Sub
    End If
    Call LogGM(UserList(UserIndex).Name, "Mensaje Broadcast repetitivo:" & rdata, False)
    MensajeRepeticion = tName
    TiempoRepeticion = tInt
    IntervaloRepeticion = Prueba1
    If TiempoRepeticion = 0 Then
        Call SendData(ToAdmins, 0, 0, "||El mensaje " & MensajeRepeticion & " será repetido cada " & IntervaloRepeticion & " minutos durante tiempo indeterminado." & FONTTYPE_FENIX)
        TiempoRepeticion = -IntervaloRepeticion
    Else
        Call SendData(ToAdmins, 0, 0, "||El mensaje " & MensajeRepeticion & " será repetido cada " & IntervaloRepeticion & " minutos durante un total de " & TiempoRepeticion & " minutos." & FONTTYPE_FENIX)
        TiempoRepeticion = TiempoRepeticion - TiempoRepeticion Mod IntervaloRepeticion
    End If
    Exit Sub
End If

If UCase$(Left$(rdata, 8)) = "/BUSCAR " Then
    rdata = Right$(rdata, Len(rdata) - 8)
    For i = 1 To UBound(ObjData)
        If InStr(1, Tilde(ObjData(i).Name), Tilde(rdata)) Then
            Call SendData(ToIndex, UserIndex, 0, "||" & i & " " & ObjData(i).Name & "." & FONTTYPE_INFO)
            N = N + 1
        End If
    Next
    If N = 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||No hubo resultados de la búsqueda: " & rdata & "." & FONTTYPE_INFO)
    Else
        Call SendData(ToIndex, UserIndex, 0, "||Hubo " & N & " resultados de la busqueda: " & rdata & "." & FONTTYPE_INFO)
    End If
    Exit Sub
End If

If UCase$(Left$(rdata, 8)) = "/CUENTA " Then
    rdata = Right$(rdata, Len(rdata) - 8)
    CuentaRegresiva = val(ReadField(1, rdata, 32)) + 1
    GMCuenta = UserList(UserIndex).POS.Map
    Exit Sub
End If


If UCase$(rdata) = "/MATA" Then
    If UserList(UserIndex).flags.TargetNpc = 0 Then Exit Sub
    Call QuitarNPC(UserList(UserIndex).flags.TargetNpc)
    Call LogGM(UserList(UserIndex).Name, "/MATA " & Npclist(UserList(UserIndex).flags.TargetNpc).Name, False)
    Exit Sub
End If

If UCase$(rdata) = "/MUERE" Then
    If UserList(UserIndex).flags.TargetNpc = 0 Then Exit Sub
    Call MuereNpc(UserList(UserIndex).flags.TargetNpc, UserIndex)
    Call LogGM(UserList(UserIndex).Name, "/MUERE " & Npclist(UserList(UserIndex).flags.TargetNpc).Name, False)
    Exit Sub
End If

If UCase$(rdata) = "/IGNORAR" Then
    If UserList(UserIndex).flags.Ignorar = 1 Then
        UserList(UserIndex).flags.Ignorar = 0
        Call SendData(ToIndex, UserIndex, 0, "||Ahora las criaturas te persiguen." & FONTTYPE_INFO)
    Else
        UserList(UserIndex).flags.Ignorar = 1
        Call SendData(ToIndex, UserIndex, 0, "||Ahora las criaturas te ignoran." & FONTTYPE_INFO)
    End If
End If

If UCase$(rdata) = "/PROTEGER" Then
    TIndex = UserList(UserIndex).flags.TargetUser
    If TIndex > 0 Then
        If UserList(TIndex).flags.Privilegios > 1 Then Exit Sub
        If UserList(TIndex).flags.Protegido = 1 Then
            UserList(TIndex).flags.Protegido = 0
            Call SendData(ToIndex, UserIndex, 0, "||Desprotegiste a " & UserList(TIndex).Name & "." & FONTTYPE_INFO)
            Call SendData(ToIndex, TIndex, 0, "||" & UserList(UserIndex).Name & " te desprotegió." & FONTTYPE_FIGHT)
        Else
            UserList(TIndex).flags.Protegido = 1
            Call SendData(ToIndex, UserIndex, 0, "||Protegiste a " & UserList(TIndex).Name & "." & FONTTYPE_INFO)
            Call SendData(ToIndex, TIndex, 0, "||" & UserList(UserIndex).Name & " te protegió. No puedes atacar ni ser atacado." & FONTTYPE_FIGHT)
        End If
    End If
End If

If Left$(UCase$(rdata), 5) = "/PRO " Then
    rdata = Right$(rdata, Len(rdata) - 5)
    TIndex = NameIndex(rdata)
    If TIndex > 0 Then
        If UserList(TIndex).flags.Privilegios > 1 Then Exit Sub
        If UserList(TIndex).flags.Protegido = 1 Then
            UserList(TIndex).flags.Protegido = 0
            Call SendData(ToIndex, UserIndex, 0, "||Desprotegiste a " & UserList(TIndex).Name & "." & FONTTYPE_INFO)
            Call SendData(ToIndex, TIndex, 0, "||" & UserList(UserIndex).Name & " te desprotegió." & FONTTYPE_FIGHT)
        Else
            UserList(TIndex).flags.Protegido = 1
            Call SendData(ToIndex, UserIndex, 0, "||Protegiste a " & UserList(TIndex).Name & "." & FONTTYPE_INFO)
            Call SendData(ToIndex, TIndex, 0, "||" & UserList(UserIndex).Name & " te protegió. No puedes atacar ni ser atacado." & FONTTYPE_FIGHT)
        End If
    End If
End If

If UCase$(Left$(rdata, 8)) = "/NOMBRE " Then
    Dim NewNick As String
    rdata = Right$(rdata, Len(rdata) - 8)
    TIndex = NameIndex(ReadField(1, rdata, Asc(" ")))
    NewNick = Right$(rdata, Len(rdata) - (Len(ReadField(1, rdata, Asc(" "))) + 1))
    If Len(NewNick) = 0 Then Exit Sub
    If TIndex = 0 Then
        Call SendData(ToIndex, UserIndex, 0, "$3E")
        Exit Sub
    End If
    If ExistePersonaje(NewNick) Then
        Call SendData(ToIndex, UserIndex, 0, "||El nombre ya existe, elige otro." & FONTTYPE_INFO)
        Exit Sub
    End If
    Call ReNombrar(TIndex, NewNick)
End If

If UCase$(Left$(rdata, 5)) = "/DEST" Then
    Call LogGM(UserList(UserIndex).Name, "/DEST", False)
    rdata = Right$(rdata, Len(rdata) - 5)
    Call EraseObj(ToMap, UserIndex, UserList(UserIndex).POS.Map, 10000, UserList(UserIndex).POS.Map, UserList(UserIndex).POS.X, UserList(UserIndex).POS.Y)
    Exit Sub
End If

If UCase$(rdata) = "/MASSDEST" Then
    For Y = UserList(UserIndex).POS.Y - MinYBorder + 1 To UserList(UserIndex).POS.Y + MinYBorder - 1
        For X = UserList(UserIndex).POS.X - MinXBorder + 1 To UserList(UserIndex).POS.X + MinXBorder - 1
            If InMapBounds(X, Y) Then _
            If MapData(UserList(UserIndex).POS.Map, X, Y).OBJInfo.OBJIndex > 0 And Not ItemEsDeMapa(UserList(UserIndex).POS.Map, X, Y) Then Call EraseObj(ToMap, UserIndex, UserList(UserIndex).POS.Map, 10000, UserList(UserIndex).POS.Map, X, Y)
        Next
    Next
    Call LogGM(UserList(UserIndex).Name, "/MASSDEST", (UserList(UserIndex).flags.Privilegios = 1))
    Exit Sub
End If

If UCase$(Left$(rdata, 10)) = "/EXPLOTAR " Then
    rdata = Right$(rdata, Len(rdata) - 10)
    
    TIndex = NameIndex(rdata)
    If TIndex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||El jugador no esta online." & FONTTYPE_FENIX)
        Exit Sub
    End If
    Call TirarTodosLosItems(TIndex)
    Call SendData(ToIndex, UserIndex, 0, "||Has explotado a " & UserList(TIndex).Name & "." & FONTTYPE_FENIX)
    Call SendData(ToIndex, TIndex, 0, "||" & UserList(TIndex).Name & " te ha explotado." & FONTTYPE_FENIX)
    Exit Sub
    End If

If UCase$(Left$(rdata, 6)) = "/KILL " Then
    rdata = Right$(rdata, Len(rdata) - 6)
    TIndex = NameIndex(rdata)
    If TIndex Then
        If UserList(TIndex).flags.Privilegios < UserList(UserIndex).flags.Privilegios Then Call UserDie(TIndex)
    End If
    Exit Sub
End If

If UCase$(Left$(rdata, 12)) = "/GANOTORNEO " Then
    rdata = Right$(rdata, Len(rdata) - 12)
   
    TIndex = UserList(UserIndex).flags.TargetUser
    If TIndex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||Debes seleccionar a un jugador!" & FONTTYPE_INFO)
        Exit Sub
    End If
   
    Call SendData(ToAll, 0, 0, "||" & UserList(UserList(UserIndex).flags.TargetUser).Name & " ganó un torneo y " & rdata & " puntos de Canje" & FONTTYPE_GANOTOQUE)
    UserList(UserList(UserIndex).flags.TargetUser).Faccion.torneos = UserList(UserList(UserIndex).flags.TargetUser).Faccion.torneos + 1
    UserList(UserList(UserIndex).flags.TargetUser).flags.Canje = UserList(UserList(UserIndex).flags.TargetUser).flags.Canje + rdata
    Call LogGM(UserList(UserIndex).Name, "Gano torneo: " & UserList(TIndex).Name & " Map:" & UserList(UserIndex).POS.Map & " X:" & UserList(UserIndex).POS.X & " Y:" & UserList(UserIndex).POS.Y, False)
    Exit Sub
End If

If UCase$(Left$(rdata, 11)) = "/GANOQUEST " Then
    rdata = Right$(rdata, Len(rdata) - 11)
   
    TIndex = UserList(UserIndex).flags.TargetUser
    If TIndex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||Debes seleccionar a un jugador!" & FONTTYPE_INFO)
        Exit Sub
    End If
   
    Call SendData(ToAll, 0, 0, "||" & UserList(UserList(UserIndex).flags.TargetUser).Name & " ganó una quest y " & rdata & " puntos de Canje" & FONTTYPE_GANOTOQUE)
    UserList(UserList(UserIndex).flags.TargetUser).Faccion.Quests = UserList(UserList(UserIndex).flags.TargetUser).Faccion.Quests + 1
        UserList(UserList(UserIndex).flags.TargetUser).flags.Canje = UserList(UserList(UserIndex).flags.TargetUser).flags.Canje + rdata
    Call LogGM(UserList(UserIndex).Name, "Ganó quest: " & UserList(TIndex).Name & " Map:" & UserList(UserIndex).POS.Map & " X:" & UserList(UserIndex).POS.X & " Y:" & UserList(UserIndex).POS.Y, False)
    Exit Sub
End If

If UCase$(Left$(rdata, 13)) = "/PERDIOTORNEO" Then
    rdata = Right$(rdata, Len(rdata) - 5)
    
    TIndex = UserList(UserIndex).flags.TargetUser
    If TIndex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||Debes seleccionar a un jugador!" & FONTTYPE_INFO)
        Exit Sub
    End If
    
    UserList(UserList(UserIndex).flags.TargetUser).Faccion.torneos = UserList(UserList(UserIndex).flags.TargetUser).Faccion.torneos - 1
    
    Call LogGM(UserList(UserIndex).Name, "Restó torneo: " & UserList(TIndex).Name & " Map:" & UserList(UserIndex).POS.Map & " X:" & UserList(UserIndex).POS.X & " Y:" & UserList(UserIndex).POS.Y, False)
    Exit Sub
End If

If UCase$(Left$(rdata, 12)) = "/PERDIOQUEST" Then
    rdata = Right$(rdata, Len(rdata) - 5)
    
    TIndex = UserList(UserIndex).flags.TargetUser
    If TIndex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||Debes seleccionar a un jugador!" & FONTTYPE_INFO)
        Exit Sub
    End If
    
    UserList(UserList(UserIndex).flags.TargetUser).Faccion.Quests = UserList(UserList(UserIndex).flags.TargetUser).Faccion.Quests - 1
    Call LogGM(UserList(UserIndex).Name, "Restó quest: " & UserList(TIndex).Name & " Map:" & UserList(UserIndex).POS.Map & " X:" & UserList(UserIndex).POS.X & " Y:" & UserList(UserIndex).POS.Y, False)
    Exit Sub
End If



If UserList(UserIndex).flags.Privilegios < 3 Then Exit Sub

If Left$(UCase$(rdata), 9) = "/INDEXPJ " Then
    rdata = Right$(rdata, Len(rdata) - 9)
    If Len(rdata) = 0 Then Exit Sub
    TIndex = IndexPJ(rdata)
    If TIndex = 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||No hay un personaje llamado " & rdata & " en la base de datos." & FONTTYPE_INFO)
    Else
        Call SendData(ToIndex, UserIndex, 0, "||El IndexPJ de " & rdata & " es " & TIndex & "." & FONTTYPE_INFO)
    End If
    Exit Sub
End If

If UCase$(rdata) = "/RESTRINGIR" Then
    If Restringido Then
        Call SendData(ToAll, 0, 0, "||La restricción de GameMasters fue desactivada servidor." & FONTTYPE_FENIX)
        Call LogGM(UserList(UserIndex).Name, "Desrestringió el servidor.", False)
    Else
        Call SendData(ToAll, 0, 0, "||La restricción de GameMasters fue activada." & FONTTYPE_FENIX)
        For i = 1 To LastUser
            DoEvents
            If UserList(i).flags.UserLogged And UserList(i).flags.Privilegios = 0 And Not UserList(i).flags.PuedeDenunciar Then Call CloseSocket(i)
        Next
        Call LogGM(UserList(UserIndex).Name, "Restringió el servidor.", False)
    End If
    Restringido = Not Restringido
    Exit Sub
End If

If UCase$(Left$(rdata, 10)) = "/CAMBIARWS" Then
    Worldsaves = Right$(rdata, Len(rdata) - 11)
    Call SendData(ToIndex, UserIndex, 0, "||Worldsave modificado a: " & Worldsaves & FONTTYPE_INFO)
    Exit Sub
End If


If UCase$(Left$(rdata, 6)) = "/BANIP" Then
    Dim BanIP As String, XNick As Boolean
    
    rdata = Right$(rdata, Len(rdata) - 7)
    
    TIndex = NameIndex(rdata)
    If TIndex <= 0 Then
        XNick = False
        Call LogGM(UserList(UserIndex).Name, "/BanIP " & rdata, False)
        BanIP = rdata
    Else
        XNick = True
        Call LogGM(UserList(UserIndex).Name, "/BanIP " & UserList(TIndex).Name & " - " & UserList(TIndex).ip, False)
        BanIP = UserList(TIndex).ip
    End If
    
    
    For LoopC = 1 To BanIps.Count
        If BanIps.Item(LoopC) = BanIP Then
            Call SendData(ToIndex, UserIndex, 0, "||La IP " & BanIP & " ya se encuentra en la lista de bans." & FONTTYPE_INFO)
            Exit Sub
        End If
    Next
    
    BanIps.Add BanIP
    Call SendData(ToAdmins, UserIndex, 0, "||" & UserList(UserIndex).Name & " Baneo la IP " & BanIP & FONTTYPE_FIGHT)
    
    If XNick Then
        Call LogBan(TIndex, UserIndex, "Ban por IP desde Nick")
        
        Call SendData(ToAdmins, 0, 0, "||" & UserList(UserIndex).Name & " echo a " & UserList(TIndex).Name & "." & FONTTYPE_FIGHT)
        Call SendData(ToAdmins, 0, 0, "||" & UserList(UserIndex).Name & " Banned a " & UserList(TIndex).Name & "." & FONTTYPE_FIGHT)
        
        
        UserList(TIndex).flags.Ban = 1
        
        Call LogGM(UserList(UserIndex).Name, "Echo a " & UserList(TIndex).Name, False)
        Call LogGM(UserList(UserIndex).Name, "BAN a " & UserList(TIndex).Name, False)
        Call CloseSocket(TIndex)
    End If
    
    Exit Sub
End If


If UCase$(Left$(rdata, 8)) = "/UNBANIP" Then
    
    
    rdata = Right$(rdata, Len(rdata) - 9)
    Call LogGM(UserList(UserIndex).Name, "/UNBANIP " & rdata, False)
    
    For LoopC = 1 To BanIps.Count
        If BanIps.Item(LoopC) = rdata Then
            BanIps.Remove LoopC
            Call SendData(ToIndex, UserIndex, 0, "||La IP " & BanIP & " se ha quitado de la lista de bans." & FONTTYPE_INFO)
            Exit Sub
        End If
    Next
    
    Call SendData(ToIndex, UserIndex, 0, "||La IP " & rdata & " NO se encuentra en la lista de bans." & FONTTYPE_INFO)
    
    Exit Sub
End If

If UCase$(Left$(rdata, 9)) = "/BanMail " Then
    Dim BanMail As String, XXNick As Boolean
    
    rdata = Right$(rdata, Len(rdata) - 9)
    
    TIndex = NameIndex(rdata)
    If TIndex <= 0 Then
        XXNick = False
        Call LogGM(UserList(UserIndex).Name, "/BanMail " & rdata, False)
        BanMail = rdata
    Else
        XXNick = True
        Call LogGM(UserList(UserIndex).Name, "/BanMail " & UserList(TIndex).Name & " - " & UserList(TIndex).Email, False)
        BanMail = UserList(TIndex).Email
    End If

    
    numeromail = GetVar(App.Path & "\logs\" & "BanMail.dat", "INIT", "Mails")
    
    For LoopC = 1 To numeromail
        If GetVar(App.Path & "\logs\" & "BanMail.dat", "Mail" & numeromail, "Mail") = BanMail Then
            Call SendData(ToIndex, UserIndex, 0, "||El mail " & BanMail & " ya se encuentra en la lista de bans." & FONTTYPE_INFO)
            Exit Sub
        End If
    Next

    
    Call WriteVar(App.Path & "\logs\" & "BanMail.dat", "Mail" & numeromail + 1, "Mail", BanMail)
    If XXNick Then Call WriteVar(App.Path & "\logs\" & "BanMail.dat", "Mail" & numeromail + 1, "User", UserList(TIndex).Name)
    Call WriteVar(App.Path & "\logs\" & "BanMail.dat", "INIT", "Mails", numeromail + 1)
   
    Call SendData(ToAdmins, UserIndex, 0, "||" & UserList(UserIndex).Name & " Baneo el mail " & BanMail & FONTTYPE_FIGHT)
    
    If XXNick Then
        Call LogBan(TIndex, UserIndex, "Ban por mail desde Nick")
        
        Call SendData(ToAdmins, 0, 0, "||" & UserList(UserIndex).Name & " echo a " & UserList(TIndex).Name & "." & FONTTYPE_FIGHT)
        Call SendData(ToAdmins, 0, 0, "||" & UserList(UserIndex).Name & " Banned a " & UserList(TIndex).Name & "." & FONTTYPE_FIGHT)
        
        
        UserList(TIndex).flags.Ban = 1
        
        Call LogGM(UserList(UserIndex).Name, "Echo a " & UserList(TIndex).Name, False)
        Call LogGM(UserList(UserIndex).Name, "BAN a " & UserList(TIndex).Name, False)
        Call CloseSocket(TIndex)
    End If
    
    Exit Sub
End If


If UCase$(Left$(rdata, 11)) = "/UNBanMail " Then
    
    numeromail = GetVar(App.Path & "\logs\" & "BanMail.dat", "INIT", "Mails")

    
    rdata = Right$(rdata, Len(rdata) - 11)
    Call LogGM(UserList(UserIndex).Name, "/UNBanMail " & rdata, False)
    
    For LoopC = 1 To numeromail
        If GetVar(App.Path & "\logs\" & "BanMail.dat", "Mail" & numeromail, "Mail") = rdata Then
            Call WriteVar(App.Path & "\logs\" & "BanMail.dat", "Mail" & numeromail, "Mail", "Desbaneado por " & UserList(UserIndex).Name)
            Call SendData(ToIndex, UserIndex, 0, "||El mail " & rdata & " se ha quitado de la lista de bans." & FONTTYPE_INFO)
            Exit Sub
        End If
    Next
    
    Call SendData(ToIndex, UserIndex, 0, "||El mail " & rdata & " NO se encuentra en la lista de bans." & FONTTYPE_INFO)
    
    Exit Sub
End If

'' con esto cancelamos el torneo
If UCase(rdata) = "/CANCELAR" Then
Call Rondas_Cancela
Exit Sub
End If
 
If UCase$(Left$(rdata, 9)) = "/SATUROS " Then
rdata = Right$(rdata, Len(rdata) - 9)
If (Torneo_Activo) Then
 Call SendData(ToIndex, UserIndex, 0, "||Ya hay un torneo!" & FONTTYPE_INFO)
 Exit Sub
End If
Dim torneos As Integer
torneos = CInt(rdata)
CantAuto = (2 ^ rdata)
If (torneos > 0 And torneos < 6) Then Call Torneos_Inicia(UserIndex, torneos)
End If

If UCase$(Left$(rdata, 3)) = "/CT" Then
    
    rdata = Right$(rdata, Len(rdata) - 4)
    Call LogGM(UserList(UserIndex).Name, "/CT: " & rdata, False)
    mapa = ReadField(1, rdata, 32)
    X = ReadField(2, rdata, 32)
    Y = ReadField(3, rdata, 32)
    
    If MapData(UserList(UserIndex).POS.Map, UserList(UserIndex).POS.X, UserList(UserIndex).POS.Y - 1).OBJInfo.OBJIndex Then
        Exit Sub
    End If
    If MapData(UserList(UserIndex).POS.Map, UserList(UserIndex).POS.X, UserList(UserIndex).POS.Y - 1).TileExit.Map Then
        Exit Sub
    End If
    If Not MapaValido(mapa) Or Not InMapBounds(X, Y) Then Exit Sub
    
    Dim ET As Obj
    ET.Amount = 1
    ET.OBJIndex = Teleport
    
    Call MakeObj(ToMap, 0, UserList(UserIndex).POS.Map, ET, UserList(UserIndex).POS.Map, UserList(UserIndex).POS.X, UserList(UserIndex).POS.Y - 1)
    MapData(UserList(UserIndex).POS.Map, UserList(UserIndex).POS.X, UserList(UserIndex).POS.Y - 1).TileExit.Map = mapa
    MapData(UserList(UserIndex).POS.Map, UserList(UserIndex).POS.X, UserList(UserIndex).POS.Y - 1).TileExit.X = X
    MapData(UserList(UserIndex).POS.Map, UserList(UserIndex).POS.X, UserList(UserIndex).POS.Y - 1).TileExit.Y = Y
    
    Exit Sub
End If



If UCase$(Left$(rdata, 3)) = "/DT" Then
    
    Call LogGM(UserList(UserIndex).Name, "/DT", False)
    
    mapa = UserList(UserIndex).flags.TargetMap
    X = UserList(UserIndex).flags.TargetX
    Y = UserList(UserIndex).flags.TargetY
    
    If ObjData(MapData(mapa, X, Y).OBJInfo.OBJIndex).ObjType = OBJTYPE_TELEPORT And _
        MapData(mapa, X, Y).TileExit.Map Then
        Call EraseObj(ToMap, 0, mapa, MapData(mapa, X, Y).OBJInfo.Amount, mapa, X, Y)
        MapData(mapa, X, Y).TileExit.Map = 0
        MapData(mapa, X, Y).TileExit.X = 0
        MapData(mapa, X, Y).TileExit.Y = 0
    End If
    
    Exit Sub
End If


If UCase$(Left$(rdata, 5)) = "/BLOQ" Then
    Call LogGM(UserList(UserIndex).Name, "/BLOQ", False)
    rdata = Right$(rdata, Len(rdata) - 5)
    If MapData(UserList(UserIndex).POS.Map, UserList(UserIndex).POS.X, UserList(UserIndex).POS.Y).Blocked = 0 Then
        MapData(UserList(UserIndex).POS.Map, UserList(UserIndex).POS.X, UserList(UserIndex).POS.Y).Blocked = 1
        Call Bloquear(ToMap, UserIndex, UserList(UserIndex).POS.Map, UserList(UserIndex).POS.Map, UserList(UserIndex).POS.X, UserList(UserIndex).POS.Y, 1)
    Else
        MapData(UserList(UserIndex).POS.Map, UserList(UserIndex).POS.X, UserList(UserIndex).POS.Y).Blocked = 0
        Call Bloquear(ToMap, UserIndex, UserList(UserIndex).POS.Map, UserList(UserIndex).POS.Map, UserList(UserIndex).POS.X, UserList(UserIndex).POS.Y, 0)
    End If
    Exit Sub
End If


If UCase$(rdata) = "/MASSKILL" Then
    For Y = UserList(UserIndex).POS.Y - MinYBorder + 1 To UserList(UserIndex).POS.Y + MinYBorder - 1
            For X = UserList(UserIndex).POS.X - MinXBorder + 1 To UserList(UserIndex).POS.X + MinXBorder - 1
                If X > 0 And Y > 0 And X < 101 And Y < 101 Then _
                    If MapData(UserList(UserIndex).POS.Map, X, Y).NpcIndex Then Call QuitarNPC(MapData(UserList(UserIndex).POS.Map, X, Y).NpcIndex)
            Next
    Next
    Call LogGM(UserList(UserIndex).Name, "/MASSKILL", False)
    Exit Sub
End If

If UCase$(Left$(rdata, 6)) = "/SMSG " Then
    rdata = Right$(rdata, Len(rdata) - 6)
    Call LogGM(UserList(UserIndex).Name, "Mensaje de sistema:" & rdata, False)
    Call SendData(ToAll, 0, 0, "!!" & rdata & ENDC)
    Exit Sub
End If

If UCase$(Left$(rdata, 7)) = "/ACT2V2" Then
    If OPCDuelos.ACT = True Then
        OPCDuelos.ACT = False
        frmMain.TimerRetos2v2.Enabled = False '> CUANDO CREEN EL TIMER, CAMBIENLEN EL NOMBRE.
        Call SendData(ToAll, 0, 0, "||Retos 2v2 Desactivados" & FONTTYPE_FENIX)
    Else
        OPCDuelos.ACT = True
        Call SendData(ToAll, 0, 0, "||Retos 2v2 Activados" & FONTTYPE_FENIX)
    End If
    Exit Sub
End If

If UCase$(Left$(rdata, 9)) = "/ACT1V1" Then
    If Actretos = True Then
        Actretos = False
        Call SendData(ToAll, 0, 0, "||Retos 1 vs 1 desactivados" & FONTTYPE_INFO)
    Else
        Call SendData(ToAll, 0, 0, "||Retos 1 vs 1 activados" & FONTTYPE_INFO)
        Actretos = True
    End If
    Exit Sub
End If


If UCase$(Left$(rdata, 5)) = "/ACC " Then
   rdata = val(Right$(rdata, Len(rdata) - 5))
   NumNPC = val(GetVar(App.Path & "\Dat\NPCs-HOSTILES.dat", "INIT", "NumNPCs")) + 500
   If rdata < 0 Or rdata > NumNPC Then
       Call SendData(ToIndex, UserIndex, 0, "||La criatura no existe." & FONTTYPE_INFO)

Else
   Call SpawnNpc(val(rdata), UserList(UserIndex).POS, True, False)


   End If
   Exit Sub
End If


If UCase$(Left$(rdata, 6)) = "/RACC " Then
   rdata = val(Right$(rdata, Len(rdata) - 6))
      NumNPC = val(GetVar(App.Path & "\Dat\NPCs-HOSTILES.dat", "INIT", "NumNPCs")) + 500
   If rdata < 0 Or rdata > NumNPC Then
    Call SendData(ToIndex, UserIndex, 0, "||La criatura no existe." & FONTTYPE_INFO)
Else
   Call SpawnNpc(val(rdata), UserList(UserIndex).POS, True, True)
   End If
   Exit Sub
End If

If UCase$(rdata) = "/NAVE" Then
    If UserList(UserIndex).flags.Navegando Then
        UserList(UserIndex).flags.Navegando = 0
    Else
        UserList(UserIndex).flags.Navegando = 1
    End If
    Exit Sub
End If

If UCase$(rdata) = "/APAGAR" Then
    Call LogMain(" Server apagado por " & UserList(UserIndex).Name & ".")
    Call ApagarSistema
    End
End If

If UCase$(rdata) = "/REINICIAR2" Then
    Call LogMain(" Server apagado especial 2 por " & UserList(UserIndex).Name & ".")
    ShellExecute frmMain.hwnd, "open", App.Path & "/fenixao2.exe", "", "", 1
    Call ApagarSistema
    Exit Sub
End If

If UCase$(rdata) = "/REINICIAR1" Then
    Call LogMain(" Server apagado especial 1 por " & UserList(UserIndex).Name & ".")
    ShellExecute frmMain.hwnd, "open", App.Path & "/fenixao.exe", "", "", 1
    Call ApagarSistema
    Exit Sub
End If

If UCase$(rdata) = "/INTERVALOS" Then
    Call SendData(ToIndex, UserIndex, 0, "||Golpe-Golpe: " & IntervaloUserPuedeAtacar & " segundos." & FONTTYPE_INFO)
    Call SendData(ToIndex, UserIndex, 0, "||Golpe-Hechizo: " & IntervaloUserPuedeGolpeHechi & " segundos." & FONTTYPE_INFO)
    Call SendData(ToIndex, UserIndex, 0, "||Hechizo-Hechizo: " & IntervaloUserPuedeCastear & " segundos." & FONTTYPE_INFO)
    Call SendData(ToIndex, UserIndex, 0, "||Hechizo-Golpe: " & IntervaloUserPuedeHechiGolpe & " segundos." & FONTTYPE_INFO)
    Call SendData(ToIndex, UserIndex, 0, "||Arco-Arco: " & IntervaloUserFlechas & " segundos." & FONTTYPE_INFO)
    Exit Sub
End If

If UCase$(Left$(rdata, 6)) = "/MODS " Then
    Dim PreInt As Single
    rdata = Right$(rdata, Len(rdata) - 6)
    TIndex = ClaseIndex(ReadField(1, rdata, 64))
    If TIndex = 0 Then Exit Sub
    tInt = ReadField(2, rdata, 64)
    If tInt < 1 Or tInt > 6 Then Exit Sub
    Arg5 = ReadField(3, rdata, 64)
    If Arg5 < 40 Or Arg5 > 125 Then Exit Sub
    PreInt = Mods(tInt, TIndex)
    Mods(tInt, TIndex) = Arg5 / 100
    Call SendData(ToAdmins, 0, 0, "||El modificador n° " & tInt & " de la clase " & ListaClases(TIndex) & " fue cambiado de " & PreInt & " a " & Mods(tInt, TIndex) & "." & FONTTYPE_FIGHT)
    Call SaveMod(tInt, TIndex)
    Exit Sub
End If

If UCase$(Left$(rdata, 4)) = "/INT" Then
    rdata = Right$(rdata, Len(rdata) - 4)
    
    Select Case UCase$(Left$(rdata, 2))
        Case "GG"
            rdata = Right$(rdata, Len(rdata) - 3)
            PreInt = IntervaloUserPuedeAtacar
            IntervaloUserPuedeAtacar = val(rdata) / 10
            Call SendData(ToAdmins, 0, 0, "||El intervalo Golpe-Golpe fue cambiado de " & PreInt & " a " & IntervaloUserPuedeAtacar & " segundos." & FONTTYPE_INFO)
            Call SendData(ToAll, 0, 0, "INTA" & IntervaloUserPuedeAtacar * 10)
            Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedeAtacar", IntervaloUserPuedeAtacar * 10)
        Case "GH"
            rdata = Right$(rdata, Len(rdata) - 3)
            PreInt = IntervaloUserPuedeGolpeHechi
            IntervaloUserPuedeGolpeHechi = val(rdata) / 10
            Call SendData(ToAdmins, 0, 0, "||El intervalo Golpe-Hechizo fue cambiado de " & PreInt & " a " & IntervaloUserPuedeGolpeHechi & " segundos." & FONTTYPE_INFO)
            Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedeGolpeHechi", IntervaloUserPuedeGolpeHechi * 10)
        Case "HH"
            rdata = Right$(rdata, Len(rdata) - 3)
            PreInt = IntervaloUserPuedeCastear
            IntervaloUserPuedeCastear = val(rdata) / 10
            Call SendData(ToAdmins, 0, 0, "||El intervalo Hechizo-Hechizo fue cambiado de " & PreInt & " a " & IntervaloUserPuedeCastear & " segundos." & FONTTYPE_INFO)
            Call SendData(ToAll, 0, 0, "INTS" & IntervaloUserPuedeCastear * 10)
            Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloLanzaHechizo", IntervaloUserPuedeCastear * 10)
        Case "HG"
            rdata = Right$(rdata, Len(rdata) - 3)
            PreInt = IntervaloUserPuedeHechiGolpe
            IntervaloUserPuedeHechiGolpe = val(rdata) / 10
            Call SendData(ToAdmins, 0, 0, "||El intervalo Hechizo-Golpe fue cambiado de " & PreInt & " a " & IntervaloUserPuedeHechiGolpe & " segundos." & FONTTYPE_INFO)
            Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedeHechiGolpe", IntervaloUserPuedeHechiGolpe * 10)
        Case "AA"
            rdata = Right$(rdata, Len(rdata) - 2)
            PreInt = IntervaloUserFlechas
            IntervaloUserFlechas = val(rdata) / 10
            Call SendData(ToAdmins, 0, 0, "||El intervalo de flechas fue cambiado de " & PreInt & " a " & IntervaloUserFlechas & " segundos." & FONTTYPE_INFO)
            Call SendData(ToIndex, UserIndex, 0, "INTF" & IntervaloUserFlechas * 10)
            
            Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserFlechas", IntervaloUserFlechas * 10)
        Case "SH"
            rdata = Right$(rdata, Len(rdata) - 2)
            PreInt = IntervaloUserSH
            IntervaloUserSH = val(rdata)
            Call SendData(ToAdmins, 0, 0, "||Intervalo de SH cambiado de " & PreInt & " a " & IntervaloUserSH & " segundos de tardanza." & FONTTYPE_INFO)
            Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserSH", str(IntervaloUserSH))
            
    End Select
End If


If UCase$(rdata) = "/DIE" Then
    Call UserDie(UserIndex)
    Exit Sub
End If

If UCase$(rdata) = "/DATS" Then
    Call CargarHechizos
    Call LoadOBJData
    Call DescargaNpcsDat
    Call CargaNpcsDat
    Exit Sub
End If

If UCase$(Left$(rdata, 6)) = "/ITEM " Then
    rdata = Right$(rdata, Len(rdata) - 6)
    ET.OBJIndex = val(ReadField(1, rdata, Asc(" ")))
    ET.Amount = val(ReadField(2, rdata, Asc(" ")))
    If ET.Amount <= 0 Then ET.Amount = 1
    If ET.OBJIndex < 1 Or ET.OBJIndex > NumObjDatas Then Exit Sub
    If ET.Amount > MAX_INVENTORY_OBJS Then Exit Sub
    If Not MeterItemEnInventario(UserIndex, ET) Then Call TirarItemAlPiso(UserList(UserIndex).POS, ET)
    Call LogGM(UserList(UserIndex).Name, "Creo objeto:" & ObjData(ET.OBJIndex).Name & " (" & ET.Amount & ")", False)
    Exit Sub
End If

If UCase$(Left$(rdata, 7)) = "/NOMANA" Then
    rdata = Right$(rdata, Len(rdata) - 7)
    UserList(UserIndex).Stats.MinMAN = 0
    Call SendUserMANA(UserIndex)
    Exit Sub
End If

If UCase$(rdata) = "/MODOQUEST" Then
    ModoQuest = Not ModoQuest
    If ModoQuest Then
        Call SendData(ToAll, 0, 0, "||Modo Quest activado." & FONTTYPE_FENIX)
        Call SendData(ToAll, 0, 0, "||Los neutrales pueden poner /MERCENARIO ALIANZA o /MERCENARIO LORD THEK para enlistarse en alguna facción temporalmente durante la quest." & FONTTYPE_FENIX)
        Call SendData(ToAll, 0, 0, "||Al morir puedes poner /HOGAR y serás teletransportado a Ullathorpe." & FONTTYPE_FENIX)
    Else
        Call SendData(ToAdmins, 0, 0, "||" & UserList(UserIndex).Name & " desactivó el modo quest." & FONTTYPE_FENIX)
        Call DesactivarMercenarios
    End If
    Exit Sub
End If

If UCase$(Left$(rdata, 5)) = "/MOD " Then
    Call LogGM(UserList(UserIndex).Name, rdata, False)
    rdata = Right$(rdata, Len(rdata) - 5)
    TIndex = NameIndex(ReadField(1, rdata, 32))
    Arg1 = ReadField(2, rdata, 32)
    Arg2 = ReadField(3, rdata, 32)
    arg3 = ReadField(4, rdata, 32)
    Arg4 = ReadField(5, rdata, 32)
    If TIndex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "1A")
        Exit Sub
    End If
    If UserList(TIndex).flags.Privilegios > 2 And UserIndex <> TIndex Then Exit Sub
    
    Select Case UCase$(Arg1)
        Case "RAZA"
            If val(Arg2) < 6 Then
                UserList(TIndex).Raza = val(Arg2)
                Call DarCuerpoDesnudo(TIndex)
                Call ChangeUserChar(ToMap, 0, UserList(UserIndex).POS.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
            End If
        Case "JER"
            UserList(UserIndex).Faccion.Jerarquia = 0
        Case "BANDO"
            If val(Arg2) < 3 Then
                If val(Arg2) > 0 Then Call SendData(ToIndex, TIndex, 0, Mensajes(val(Arg2), 10))
                UserList(TIndex).Faccion.Bando = val(Arg2)
                UserList(TIndex).Faccion.BandoOriginal = val(Arg2)
                If Not PuedeFaccion(TIndex) Then Call SendData(ToIndex, TIndex, 0, "SUFA0")
                Call UpdateUserChar(TIndex)
                If val(Arg2) = 0 Then UserList(TIndex).Faccion.Jerarquia = 0
            End If
        Case "SKI"
            If val(Arg2) >= 0 And val(Arg2) <= 100 Then
                For i = 1 To NUMSKILLS
                    UserList(TIndex).Stats.UserSkills(i) = val(Arg2)
                Next
            End If
        Case "CLASE"
            i = ClaseIndex(Arg2)
            If i = 0 Then Exit Sub
            UserList(TIndex).Clase = i
            UserList(TIndex).Recompensas(1) = 0
            UserList(TIndex).Recompensas(2) = 0
            UserList(TIndex).Recompensas(3) = 0
            Call SendData(ToIndex, TIndex, 0, "||Ahora eres " & ListaClases(i) & "." & FONTTYPE_INFO)
            If PuedeRecompensa(UserIndex) Then
                Call SendData(ToIndex, UserIndex, 0, "SURE1")
            Else: Call SendData(ToIndex, UserIndex, 0, "SURE0")
            End If
            If PuedeSubirClase(UserIndex) Then
                Call SendData(ToIndex, UserIndex, 0, "SUCL1")
            Else: Call SendData(ToIndex, UserIndex, 0, "SUCL0")
            End If
        
        Case "ORO"
            If val(Arg2) > 10000000 Then Arg2 = 10000000
            UserList(TIndex).Stats.GLD = val(Arg2)
            Call SendUserORO(TIndex)
        Case "EXP"
            If val(Arg2) > 10000000 Then Arg2 = 10000000
            UserList(TIndex).Stats.Exp = val(Arg2)
            Call CheckUserLevel(TIndex)
            Call SendUserEXP(TIndex)
        Case "MEX"
            If val(Arg2) > 10000000 Then Arg2 = 10000000
            UserList(TIndex).Stats.Exp = UserList(TIndex).Stats.Exp + val(Arg2)
            Call CheckUserLevel(TIndex)
            Call SendUserEXP(TIndex)
        Case "BODY"
            Call ChangeUserBody(ToMap, 0, UserList(TIndex).POS.Map, TIndex, val(Arg2))
        Case "HEAD"
            Call ChangeUserHead(ToMap, 0, UserList(TIndex).POS.Map, TIndex, val(Arg2))
            UserList(TIndex).OrigChar.Head = val(Arg2)
        Case "PHEAD"
            UserList(TIndex).OrigChar.Head = val(Arg2)
            Call ChangeUserHead(ToMap, 0, UserList(TIndex).POS.Map, TIndex, val(Arg2))
        Case "TOR"
            UserList(TIndex).Faccion.torneos = val(Arg2)
        Case "QUE"
            UserList(TIndex).Faccion.Quests = val(Arg2)
        Case "NEU"
            UserList(TIndex).Faccion.Matados(Neutral) = val(Arg2)
        Case "CRI"
            UserList(TIndex).Faccion.Matados(Caos) = val(Arg2)
        Case "CIU"
            UserList(TIndex).Faccion.Matados(Real) = val(Arg2)
        Case "HP"
            If val(Arg2) > 999 Then Exit Sub
            UserList(TIndex).Stats.MaxHP = val(Arg2)
            Call SendUserMAXHP(UserIndex)
        Case "MAN"
            If val(Arg2) > 2200 + 800 * Buleano(UserList(TIndex).Clase = MAGO And UserList(TIndex).Recompensas(2) = 2) Then Exit Sub
            UserList(TIndex).Stats.MaxMAN = val(Arg2)
            Call SendUserMAXMANA(UserIndex)
        Case "STA"
            If val(Arg2) > 999 Then Exit Sub
            UserList(TIndex).Stats.MaxSta = val(Arg2)
        Case "HAM"
            UserList(TIndex).Stats.MinHam = val(Arg2)
        Case "SED"
            UserList(TIndex).Stats.MinAGU = val(Arg2)
        Case "ATF"
            If val(Arg2) > 21 Or val(Arg2) < 6 Then Exit Sub
            UserList(TIndex).Stats.UserAtributos(fuerza) = val(Arg2)
            UserList(TIndex).Stats.UserAtributosBackUP(fuerza) = val(Arg2)
            Call UpdateFuerzaYAg(TIndex)
        Case "ATI"
            If val(Arg2) > 21 Or val(Arg2) < 6 Then Exit Sub
            UserList(TIndex).Stats.UserAtributos(Inteligencia) = val(Arg2)
            UserList(TIndex).Stats.UserAtributosBackUP(Inteligencia) = val(Arg2)
        Case "ATA"
            If val(Arg2) > 21 Or val(Arg2) < 6 Then Exit Sub
            UserList(TIndex).Stats.UserAtributos(Agilidad) = val(Arg2)
            UserList(TIndex).Stats.UserAtributosBackUP(Agilidad) = val(Arg2)
            Call UpdateFuerzaYAg(TIndex)
            
        Case "RETO"
            UserList(TIndex).flags.Reto = val(Arg2)
        Case "ATC"
            If val(Arg2) > 21 Or val(Arg2) < 6 Then Exit Sub
            UserList(TIndex).Stats.UserAtributos(Carisma) = val(Arg2)
            UserList(TIndex).Stats.UserAtributosBackUP(Carisma) = val(Arg2)
        Case "ATV"
            If val(Arg2) > 21 Or val(Arg2) < 6 Then Exit Sub
            UserList(TIndex).Stats.UserAtributos(Constitucion) = val(Arg2)
            UserList(TIndex).Stats.UserAtributosBackUP(Constitucion) = val(Arg2)
        Case "LEVEL"
            If val(Arg2) < 1 Or val(Arg2) > STAT_MAXELV Then Exit Sub
            UserList(TIndex).Stats.ELV = val(Arg2)
            UserList(TIndex).Stats.ELU = ELUs(UserList(TIndex).Stats.ELV)
            Call SendData(ToIndex, TIndex, 0, "5O" & UserList(TIndex).Stats.ELV & "," & UserList(TIndex).Stats.ELU)
            If PuedeRecompensa(UserIndex) Then
                Call SendData(ToIndex, UserIndex, 0, "SURE1")
            Else: Call SendData(ToIndex, UserIndex, 0, "SURE0")
            End If
            If PuedeSubirClase(UserIndex) Then
                Call SendData(ToIndex, UserIndex, 0, "SUCL1")
            Else: Call SendData(ToIndex, UserIndex, 0, "SUCL0")
            End If
        Case Else
            Call SendData(ToIndex, UserIndex, 0, "||Comando inexistente." & FONTTYPE_INFO)
    End Select

    Exit Sub
End If

If UCase$(Left$(rdata, 9)) = "/DOBACKUP" Then
    Call DoBackUp
    Exit Sub
End If

If UCase$(Left$(rdata, 10)) = "/DOBACKUPL" Then
    Call DoBackUp(True)
    Exit Sub
End If

If UCase$(Left$(rdata, 7)) = "/GRABAR" Then
    Call GuardarUsuarios
    Exit Sub
End If

If UCase$(Left$(rdata, 9)) = "/PAUSA" Then
    
    If haciendoBK Then Exit Sub
    
    Enpausa = Not Enpausa
    
    If Enpausa Then
        Call SendData(ToAll, 0, 0, "TL" & 197)
        Call SendData(ToAll, 0, 0, "||Servidor> El mundo ha sido detenido." & FONTTYPE_INFO)
        Call SendData(ToAll, 0, 0, "BKW")
        Call SendData(ToAll, 0, 0, "TM" & "0")
    Else
        Call SendData(ToAll, 0, 0, "TL")
        Call SendData(ToAll, 0, 0, "||Servidor> Juego reanudado." & FONTTYPE_INFO)
        Call SendData(ToAll, 0, 0, "BKW")
        Call SendData(ToIndex, UserIndex, 0, "TM" & MapInfo(UserList(UserIndex).POS.Map).Music)
    End If
Exit Sub
End If

If UCase$(Left$(rdata, 9)) = "/SHOW INT" Then
    Call frmMain.mnuMostrar_Click
    Exit Sub
End If

If UCase$(rdata) = "/LLUVIA" Then
    Lloviendo = Not Lloviendo
    Call SendData(ToAll, 0, 0, "LLU")
    Exit Sub
End If

If UCase$(rdata) = "/PASSDAY" Then
    Call DayElapsed
    Exit Sub
End If


Exit Sub

ErrorHandler:
 Call LogError("HandleData. CadOri:" & CadenaOriginal & " Nom:" & UserList(UserIndex).Name & " UI:" & UserIndex & " N: " & Err.Number & " D: " & Err.Description)
 Call Cerrar_Usuario(UserIndex)

End Sub
