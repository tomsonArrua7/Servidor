Attribute VB_Name = "SistemaCombate"
'================================
'>>>>> WWW.FADICTOS.COM.AR <<<<<<
'================================
'Matute - matius_xd07@hotmail.com

Option Explicit

Public Declare Function Minimo Lib "aolib.dll" (ByVal A As Long, ByVal B As Long) As Long
Public Declare Function Maximo Lib "aolib.dll" (ByVal A As Long, ByVal B As Long) As Long
Public Declare Function PoderAtaqueWresterling Lib "aolib.dll" (ByVal Skill As Byte, ByVal Agilidad As Integer, Clase As Byte, ByVal Nivel As Byte) As Integer
Public Declare Function SD Lib "aolib.dll" (ByVal N As Integer) As Integer
Public Declare Function SDM Lib "aolib.dll" (ByVal N As Integer) As Integer
Public Declare Function Complex Lib "aolib.dll" (ByVal N As Integer) As Integer
Public Declare Function RandomNumber Lib "aolib.dll" (ByVal MIN As Long, ByVal MAX As Long) As Long

Public Const EVASION = 1
Public Const CUERPOACUERPO = 2
Public Const CONARCOS = 3
Public Const EVAESCUDO = 4
Public Const DANOCUERPOACUERPO = 5
Public Const DANOCONARCOS = 6

Public Mods(1 To 6, 1 To NUMCLASES) As Single
Public Const MAXDISTANCIAARCO = 12
Public Sub CargarMods()
Dim i As Byte, j As Integer
Dim file As String

file = DatPath & "Mods.dat"

For i = 1 To NUMCLASES
    If Len(ListaClases(i)) > 0 Then
        For j = 1 To UBound(Mods, 1)
            Mods(j, i) = Int(GetVar(file, ListaClases(i), "Mod" & j)) / 100
        Next
    End If
Next

End Sub
Public Sub SaveMod(A As Integer, B As Integer)

Call WriteVar(DatPath & "Mods.dat", ListaClases(B), "Mod" & A, str(Mods(A, B) * 100))

End Sub

Public Function PoderAtaqueProyectil(UserIndex As Integer) As Integer

Select Case UserList(UserIndex).Stats.UserSkills(Proyectiles)
    Case Is < 31
        PoderAtaqueProyectil = UserList(UserIndex).Stats.UserSkills(Proyectiles) * Mods(CONARCOS, UserList(UserIndex).Clase)
    Case Is < 61
        PoderAtaqueProyectil = (UserList(UserIndex).Stats.UserSkills(Proyectiles) + UserList(UserIndex).Stats.UserAtributos(Agilidad)) * Mods(CONARCOS, UserList(UserIndex).Clase)
    Case Is < 91
        PoderAtaqueProyectil = (UserList(UserIndex).Stats.UserSkills(Proyectiles) + 2 * UserList(UserIndex).Stats.UserAtributos(Agilidad)) * Mods(CONARCOS, UserList(UserIndex).Clase)
    Case Else
        PoderAtaqueProyectil = (UserList(UserIndex).Stats.UserSkills(Proyectiles) + 3 * UserList(UserIndex).Stats.UserAtributos(Agilidad)) * Mods(CONARCOS, UserList(UserIndex).Clase)
End Select

PoderAtaqueProyectil = (PoderAtaqueProyectil + (2.5 * Maximo(UserList(UserIndex).Stats.ELV - 12, 0)))

End Function
Public Function PoderAtaqueArma(UserIndex As Integer) As Integer

Select Case UserList(UserIndex).Stats.UserSkills(Armas)
    Case Is < 31
        PoderAtaqueArma = UserList(UserIndex).Stats.UserSkills(Armas) * Mods(CUERPOACUERPO, UserList(UserIndex).Clase)
    Case Is < 61
        PoderAtaqueArma = (UserList(UserIndex).Stats.UserSkills(Armas) + UserList(UserIndex).Stats.UserAtributos(Agilidad)) * Mods(CUERPOACUERPO, UserList(UserIndex).Clase)
    Case Is < 91
        PoderAtaqueArma = (UserList(UserIndex).Stats.UserSkills(Armas) + 2 * UserList(UserIndex).Stats.UserAtributos(Agilidad)) * Mods(CUERPOACUERPO, UserList(UserIndex).Clase)
    Case Else
        PoderAtaqueArma = (UserList(UserIndex).Stats.UserSkills(Armas) + 3 * UserList(UserIndex).Stats.UserAtributos(Agilidad)) * Mods(CUERPOACUERPO, UserList(UserIndex).Clase)
End Select

PoderAtaqueArma = PoderAtaqueArma + 2.5 * Maximo(UserList(UserIndex).Stats.ELV - 12, 0)

End Function
Public Function PoderEvasionEscudo(UserIndex As Integer)

PoderEvasionEscudo = UserList(UserIndex).Stats.UserSkills(Defensa) * Mods(EVAESCUDO, UserList(UserIndex).Clase) / 2

End Function
Public Function PoderEvasion(UserIndex As Integer) As Integer

Select Case UserList(UserIndex).Stats.UserSkills(Tacticas)
    Case Is < 31
        PoderEvasion = UserList(UserIndex).Stats.UserSkills(Tacticas) * Mods(EVASION, UserList(UserIndex).Clase)
    Case Is < 61
        PoderEvasion = (UserList(UserIndex).Stats.UserSkills(Tacticas) + UserList(UserIndex).Stats.UserAtributos(Agilidad)) * Mods(EVASION, UserList(UserIndex).Clase)
    Case Is < 91
        PoderEvasion = (UserList(UserIndex).Stats.UserSkills(Tacticas) + 2 * UserList(UserIndex).Stats.UserAtributos(Agilidad)) * Mods(EVASION, UserList(UserIndex).Clase)
    Case Else
        PoderEvasion = (UserList(UserIndex).Stats.UserSkills(Tacticas) + 3 * UserList(UserIndex).Stats.UserAtributos(Agilidad)) * Mods(EVASION, UserList(UserIndex).Clase)
End Select

PoderEvasion = PoderEvasion + (2.5 * Maximo(UserList(UserIndex).Stats.ELV - 12, 0))

End Function
Public Function UserImpactoNpc(UserIndex As Integer, NpcIndex As Integer) As Boolean
Dim PoderAtaque As Long
Dim Arma As Integer
Dim proyectil As Boolean
Dim ProbExito As Long

Arma = UserList(UserIndex).Invent.WeaponEqpObjIndex
If Arma = 0 Then proyectil = False Else proyectil = ObjData(Arma).proyectil = 1

If Arma = 0 Then
    PoderAtaque = PoderAtaqueWresterling(UserList(UserIndex).Stats.UserSkills(Wresterling), UserList(UserIndex).Stats.UserAtributos(Agilidad), UserList(UserIndex).Clase, UserList(UserIndex).Stats.ELV) \ 4
ElseIf proyectil Then
    PoderAtaque = (1 + 0.05 * Buleano(UserList(UserIndex).Clase = ARQUERO And UserList(UserIndex).Recompensas(3) = 1) + 0.1 * Buleano(UserList(UserIndex).Recompensas(3) = 1 And (UserList(UserIndex).Clase = GUERRERO Or UserList(UserIndex).Clase = CAZADOR))) _
    * PoderAtaqueProyectil(UserIndex)
Else
    PoderAtaque = (1 + 0.05 * Buleano(UserList(UserIndex).Clase = PALADIN And UserList(UserIndex).Recompensas(3) = 2)) _
    * PoderAtaqueArma(UserIndex)
End If

ProbExito = Maximo(10, Minimo(90, 50 + ((PoderAtaque - Npclist(NpcIndex).PoderEvasion) * 0.4)))

UserImpactoNpc = (RandomNumber(1, 100) <= ProbExito)

If UserImpactoNpc Then
    If Arma Then
       If proyectil Then
            Call SubirSkill(UserIndex, Proyectiles)
       Else: Call SubirSkill(UserIndex, Armas)
       End If
    Else
        Call SubirSkill(UserIndex, Wresterling)
    End If
End If


End Function
Public Function NpcImpacto(ByVal NpcIndex As Integer, UserIndex As Integer) As Boolean
Dim Rechazo As Boolean
Dim ProbRechazo As Long
Dim ProbExito As Long
Dim UserEvasion As Long

UserEvasion = (1 + 0.05 * Buleano(UserList(UserIndex).Recompensas(3) = 2 And (UserList(UserIndex).Clase = ARQUERO Or UserList(UserIndex).Clase = NIGROMANTE))) _
            * PoderEvasion(UserIndex)

If UserList(UserIndex).Invent.EscudoEqpObjIndex Then UserEvasion = UserEvasion + PoderEvasionEscudo(UserIndex)

ProbExito = Maximo(10, Minimo(90, 50 + ((Npclist(NpcIndex).PoderAtaque - UserEvasion) * 0.4)))

NpcImpacto = (RandomNumber(1, 100) <= ProbExito)

If UserList(UserIndex).Invent.EscudoEqpObjIndex Then
   If Not NpcImpacto Then
      ProbRechazo = Maximo(10, Minimo(90, 100 * (UserList(UserIndex).Stats.UserSkills(Defensa) / (UserList(UserIndex).Stats.UserSkills(Defensa) + UserList(UserIndex).Stats.UserSkills(Tacticas)))))
      Rechazo = (RandomNumber(1, 100) <= ProbRechazo)
      If Rechazo Then
         Call SendData(ToPCArea, UserIndex, UserList(UserIndex).POS.Map, "TW" & SND_ESCUDO)
         Call SendData(ToIndex, UserIndex, 0, "7")
         Call SubirSkill(UserIndex, Defensa, 25)
      End If
   End If
End If

End Function
Public Function CalcularDa�o(UserIndex As Integer, Optional ByVal Dragon As Boolean) As Long
Dim ModifClase As Single
Dim Da�oUsuario As Long
Dim Da�oArma As Long
Dim Da�oMaxArma As Long
Dim Arma As ObjData

Da�oUsuario = RandomNumber(UserList(UserIndex).Stats.MinHit, UserList(UserIndex).Stats.MaxHit)

If UserList(UserIndex).Invent.WeaponEqpObjIndex = 0 Then
    ModifClase = Mods(DANOCUERPOACUERPO, UserList(UserIndex).Clase)
    CalcularDa�o = Maximo(0, (UserList(UserIndex).Stats.UserAtributos(fuerza) - 15)) + Da�oUsuario * ModifClase
    Exit Function
End If

Arma = ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex)

Da�oMaxArma = Arma.MaxHit
        
If Arma.proyectil Then
    ModifClase = Mods(DANOCONARCOS, UserList(UserIndex).Clase)
    Da�oArma = RandomNumber(Arma.MinHit, Da�oMaxArma) + RandomNumber(ObjData(UserList(UserIndex).Invent.MunicionEqpObjIndex).MinHit + 10 * Buleano(UserList(UserIndex).flags.BonusFlecha) + 5 * Buleano(UserList(UserIndex).Clase = ARQUERO And UserList(UserIndex).Recompensas(3) = 2), ObjData(UserList(UserIndex).Invent.MunicionEqpObjIndex).MaxHit + 15 * Buleano(UserList(UserIndex).flags.BonusFlecha) + 3 * Buleano(UserList(UserIndex).Clase = ARQUERO And UserList(UserIndex).Recompensas(3) = 2))
Else
    ModifClase = Mods(DANOCUERPOACUERPO, UserList(UserIndex).Clase)
    If Arma.SubTipo = MATADRAGONES And Not Dragon Then
        CalcularDa�o = 1
        Exit Function
    Else
        Da�oArma = RandomNumber(Arma.MinHit, Da�oMaxArma)
    End If
End If

CalcularDa�o = (((3 * Da�oArma) + ((Da�oMaxArma / 5) * Maximo(0, (UserList(UserIndex).Stats.UserAtributos(fuerza) - 15))) + Da�oUsuario) * ModifClase)

End Function
Public Sub UserDa�oNpc(UserIndex As Integer, ByVal NpcIndex As Integer)
Dim Muere As Boolean
Dim Da�o As Long
Dim j As Integer

Da�o = CalcularDa�o(UserIndex, Npclist(NpcIndex).NPCtype = 6)

If UserList(UserIndex).flags.Navegando = 1 Then Da�o = Da�o + RandomNumber(ObjData(UserList(UserIndex).Invent.BarcoObjIndex).MinHit, ObjData(UserList(UserIndex).Invent.BarcoObjIndex).MaxHit)

Da�o = Maximo(0, Da�o - Npclist(NpcIndex).Stats.Def)

Call SendData(ToIndex, UserIndex, 0, "U2" & Da�o)

Call ExperienciaPorGolpe(UserIndex, NpcIndex, CInt(Da�o))
If Da�o >= Npclist(NpcIndex).Stats.MinHP Then Muere = True
Call VerNPCMuere(NpcIndex, Da�o, UserIndex)

If Not Muere Then
    If PuedeApu�alar(UserIndex) Then
       Call DoApu�alar(UserIndex, NpcIndex, 0, CInt(Da�o))
       Call SubirSkill(UserIndex, Apu�alar)
    End If
End If

End Sub
Public Sub NpcDa�o(ByVal NpcIndex As Integer, UserIndex As Integer)
Dim Da�o As Integer, lugar As Integer, absorbido As Integer, npcfile As String
Dim antda�o As Integer, defbarco As Integer
Dim Obj As ObjData
Dim Obj2 As ObjData

Da�o = RandomNumber(Npclist(NpcIndex).Stats.MinHit, Npclist(NpcIndex).Stats.MaxHit)
antda�o = Da�o

If UserList(UserIndex).flags.Navegando = 1 Then
    Obj = ObjData(UserList(UserIndex).Invent.BarcoObjIndex)
    defbarco = RandomNumber(Obj.MinDef, Obj.MaxDef)
End If

If UserList(UserIndex).flags.Montado = 1 Then
     defbarco = defbarco + UserList(UserIndex).Caballos.Agi(UserList(UserIndex).flags.CaballoMontado)
End If

lugar = RandomNumber(1, 6)

Select Case lugar
  Case bCabeza
        
        If UserList(UserIndex).Invent.CascoEqpObjIndex Then
            Obj = ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex)
            If Obj.Gorro = 0 Then absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
        End If
        
  Case Else
        
        If UserList(UserIndex).Invent.ArmourEqpObjIndex Then
           Obj = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex)
           absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
        End If
        
        If UserList(UserIndex).Invent.EscudoEqpObjIndex Then
           Obj2 = ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex)
            absorbido = absorbido + RandomNumber(Obj2.MinDef, Obj2.MaxDef)
        End If
        
End Select

absorbido = absorbido + defbarco + 2 * Buleano(UserList(UserIndex).Clase = GUERRERO And UserList(UserIndex).Recompensas(2) = 2)

Da�o = Maximo(1, Da�o - absorbido)

Call SendData(ToIndex, UserIndex, 0, "N2" & lugar & "," & Da�o)

If UserList(UserIndex).flags.Privilegios = 0 And Not UserList(UserIndex).flags.Quest Then UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP - Da�o

If UserList(UserIndex).Stats.MinHP <= 0 Then

    Call SendData(ToIndex, UserIndex, 0, "6")
    
   
    If Npclist(NpcIndex).MaestroUser Then
        Call AllFollowAmo(Npclist(NpcIndex).MaestroUser)
    Else
        
        If Npclist(NpcIndex).Stats.Alineacion = 0 Then
            Npclist(NpcIndex).Movement = Npclist(NpcIndex).flags.OldMovement
            Npclist(NpcIndex).Hostile = Npclist(NpcIndex).flags.OldHostil
            Npclist(NpcIndex).flags.AttackedBy = 0
        End If
    End If
    
    Call UserDie(UserIndex)

End If

End Sub
Public Sub CheckPets(ByVal NpcIndex As Integer, UserIndex As Integer)
Dim j As Integer

For j = 1 To MAXMASCOTAS - 17 * Buleano(Not UserList(UserIndex).flags.Quest)
    If UserList(UserIndex).MascotasIndex(j) Then
       If UserList(UserIndex).MascotasIndex(j) <> NpcIndex Then
        If Npclist(UserList(UserIndex).MascotasIndex(j)).TargetNpc = 0 Then Npclist(UserList(UserIndex).MascotasIndex(j)).TargetNpc = NpcIndex
        Npclist(UserList(UserIndex).MascotasIndex(j)).Movement = NPC_ATACA_NPC
       End If
    End If
Next

End Sub
Public Sub AllFollowAmo(UserIndex As Integer)
Dim j As Integer

For j = 1 To MAXMASCOTAS - 17 * Buleano(Not UserList(UserIndex).flags.Quest)
    If UserList(UserIndex).MascotasIndex(j) Then
        Call FollowAmo(UserList(UserIndex).MascotasIndex(j))
    End If
Next

End Sub
Public Sub NpcAtacaUser(ByVal NpcIndex As Integer, UserIndex As Integer)

If Npclist(NpcIndex).AutoCurar = 1 Then Exit Sub
If Npclist(NpcIndex).Numero = 92 Then Exit Sub

If Npclist(NpcIndex).CanAttack = 1 Then
    Call CheckPets(NpcIndex, UserIndex)
    
    If Npclist(NpcIndex).Target = 0 Then Npclist(NpcIndex).Target = UserIndex
    
    If UserList(UserIndex).flags.AtacadoPorNpc = 0 And _
       UserList(UserIndex).flags.AtacadoPorUser = 0 Then UserList(UserIndex).flags.AtacadoPorNpc = NpcIndex
Else
    Exit Sub
End If

Npclist(NpcIndex).CanAttack = 0

If Npclist(NpcIndex).flags.Snd1 Then Call SendData(ToPCArea, UserIndex, UserList(UserIndex).POS.Map, "TW" & Npclist(NpcIndex).flags.Snd1)
        
If NpcImpacto(NpcIndex, UserIndex) Then
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).POS.Map, "TW" & SND_IMPACTO)
    
    If UserList(UserIndex).flags.Navegando = 0 And Not UserList(UserIndex).flags.Meditando Then Call SendData(ToPCArea, UserIndex, UserList(UserIndex).POS.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & FXSANGRE & "," & 0)

    Call NpcDa�o(NpcIndex, UserIndex)

    If Npclist(NpcIndex).Veneno = 1 Then Call NpcEnvenenarUser(UserIndex)
Else
    Call SendData(ToIndex, UserIndex, 0, "N1")
End If

Call SubirSkill(UserIndex, Tacticas)
Call SendUserHP(UserIndex)

End Sub
Function NpcImpactoNpc(ByVal Atacante As Integer, ByVal Victima As Integer) As Boolean

NpcImpactoNpc = (RandomNumber(1, 100) <= Maximo(10, Minimo(90, 50 + ((Npclist(Atacante).PoderAtaque - Npclist(Victima).PoderEvasion) * 0.4))))

End Function
Public Sub NpcDa�oNpc(ByVal Atacante As Integer, ByVal Victima As Integer)
Dim Da�o As Integer
Dim ANpc As Npc
ANpc = Npclist(Atacante)

Da�o = RandomNumber(ANpc.Stats.MinHit, ANpc.Stats.MaxHit)

If ANpc.MaestroUser Then Call ExperienciaPorGolpe(ANpc.MaestroUser, Victima, Da�o)
Call VerNPCMuere(Victima, Da�o, ANpc.MaestroUser)

If Npclist(Victima).Stats.MinHP <= 0 Then
    Call RestoreOldMovement(Atacante)
    If ANpc.MaestroUser Then Call FollowAmo(Atacante)
End If

End Sub
Public Sub NpcAtacaNpc(ByVal Atacante As Integer, ByVal Victima As Integer)

If Npclist(Atacante).CanAttack = 1 Then
    Npclist(Atacante).CanAttack = 0
    Npclist(Victima).TargetNpc = Atacante
Else: Exit Sub
End If

If Npclist(Atacante).flags.Snd1 Then Call SendData(ToNPCArea, Atacante, Npclist(Atacante).POS.Map, "TW" & Npclist(Atacante).flags.Snd1)

If NpcImpactoNpc(Atacante, Victima) Then
    
    If Npclist(Victima).flags.Snd2 Then
        Call SendData(ToNPCArea, Victima, Npclist(Victima).POS.Map, "TW" & Npclist(Victima).flags.Snd2)
    Else: Call SendData(ToNPCArea, Victima, Npclist(Victima).POS.Map, "TW" & SND_IMPACTO2)
    End If

    If Npclist(Atacante).MaestroUser Then
        Call SendData(ToNPCArea, Atacante, Npclist(Atacante).POS.Map, "TW" & SND_IMPACTO)
    Else: Call SendData(ToNPCArea, Victima, Npclist(Victima).POS.Map, "TW" & SND_IMPACTO)
    End If
    Call NpcDa�oNpc(Atacante, Victima)
    
Else
    If Npclist(Atacante).MaestroUser Then
        Call SendData(ToNPCArea, Atacante, Npclist(Atacante).POS.Map, "TW" & SOUND_SWING)
    Else
        Call SendData(ToNPCArea, Victima, Npclist(Victima).POS.Map, "TW" & SOUND_SWING)
    End If
End If

End Sub
Public Sub UsuarioAtaca(UserIndex As Integer)

If UserList(UserIndex).flags.Protegido = 1 Then
    Call SendData(ToIndex, UserIndex, 0, "||No pod�s atacar mientras est�s siendo protegido por un GM." & FONTTYPE_INFO)
    Exit Sub
ElseIf UserList(UserIndex).flags.Protegido = 2 Then
    Call SendData(ToIndex, UserIndex, 0, "||No pod�s atacar tan pronto al conectarte." & FONTTYPE_INFO)
    Exit Sub
End If

If TiempoTranscurrido(UserList(UserIndex).Counters.LastGolpe) < IntervaloUserPuedeAtacar Then Exit Sub
If TiempoTranscurrido(UserList(UserIndex).Counters.LastHechizo) < IntervaloUserPuedeHechiGolpe Then Exit Sub
If TiempoTranscurrido(UserList(UserIndex).Counters.LastFlecha) < IntervaloUserFlechas Then Exit Sub

UserList(UserIndex).Counters.LastGolpe = Timer
Call SendData(ToIndex, UserIndex, 0, "LG")

If UserList(UserIndex).flags.Oculto Then
    If Not ((UserList(UserIndex).Clase = CAZADOR Or UserList(UserIndex).Clase = ARQUERO) And UserList(UserIndex).Invent.ArmourEqpObjIndex = 360) Then
        UserList(UserIndex).flags.Oculto = 0
        UserList(UserIndex).flags.Invisible = 0
        Call SendData(ToMap, 0, UserList(UserIndex).POS.Map, ("V3" & UserList(UserIndex).Char.CharIndex & ",0"))
        Call SendData(ToIndex, UserIndex, 0, "V5")
    End If
End If

If UserList(UserIndex).Stats.MinSta >= 10 Then
    Call QuitarSta(UserIndex, RandomNumber(1, 10))
Else: Call SendData(ToIndex, UserIndex, 0, "9E")
    Exit Sub
End If

Dim AttackPos As WorldPos
AttackPos = UserList(UserIndex).POS
Call HeadtoPos(UserList(UserIndex).Char.Heading, AttackPos)

If AttackPos.X < XMinMapSize Or AttackPos.X > XMaxMapSize Or AttackPos.Y <= YMinMapSize Or AttackPos.Y > YMaxMapSize Then
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).POS.Map, "-" & UserList(UserIndex).Char.CharIndex)
    Exit Sub
End If

Dim Index As Integer
Index = MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).UserIndex

If Index Then
    Call UsuarioAtacaUsuario(UserIndex, MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).UserIndex)
    Call SendUserSTA(UserIndex)
    Call SendUserHP(MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).UserIndex)
    Exit Sub
End If

If MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).NpcIndex Then

    If Npclist(MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).NpcIndex).Attackable Then
        
        If (Npclist(MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).NpcIndex).MaestroUser > 0 And _
           MapInfo(Npclist(MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).NpcIndex).POS.Map).Pk = False) And (UserList(UserIndex).POS.Map <> 190) Then
            Call SendData(ToIndex, UserIndex, 0, "0Z")
            Exit Sub
        End If
           
        Call UsuarioAtacaNpc(UserIndex, MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).NpcIndex)

    Else
        Call SendData(ToIndex, UserIndex, 0, "NO")
    End If
    
    Call SendUserSTA(UserIndex)
    
    Exit Sub


End If

Call SendData(ToPCArea, UserIndex, UserList(UserIndex).POS.Map, "-" & UserList(UserIndex).Char.CharIndex)
Call SendUserSTA(UserIndex)

End Sub
Public Sub UsuarioAtacaNpc(UserIndex As Integer, ByVal NpcIndex As Integer)

'If Distancia(UserList(UserIndex).POS, Npclist(NpcIndex).POS) > MAXDISTANCIAARCO Then
'   Call SendData(ToIndex, UserIndex, 0, "3G")
'   Exit Sub
'End If

If (UserList(UserIndex).Faccion.Bando <> Neutral Or EsNewbie(UserIndex)) And Npclist(NpcIndex).MaestroUser Then
    If Not PuedeAtacarMascota(UserIndex, (Npclist(NpcIndex).MaestroUser)) Then Exit Sub
End If

If Npclist(NpcIndex).flags.Faccion <> Neutral Then
    If UserList(UserIndex).Faccion.Bando <> Neutral And UserList(UserIndex).Faccion.Bando = Npclist(NpcIndex).flags.Faccion Then
        Call SendData(ToIndex, UserIndex, 0, Mensajes(Npclist(NpcIndex).flags.Faccion, 19))
        Exit Sub
    ElseIf EsNewbie(UserIndex) Then
        Call SendData(ToIndex, UserIndex, 0, "%L")
        Exit Sub
    End If
End If

If UserList(UserIndex).flags.Protegido > 0 Then
    Call SendData(ToIndex, UserIndex, 0, "||No podes atacar NPC's mientr�s est�s siendo protegido." & FONTTYPE_FIGHT)
    Exit Sub
End If

Call NpcAtacado(NpcIndex, UserIndex)

If UserImpactoNpc(UserIndex, NpcIndex) Then
    If Npclist(NpcIndex).flags.Snd2 Then
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).POS.Map, "\" & UserList(UserIndex).Char.CharIndex & "," & Npclist(NpcIndex).flags.Snd2)
    Else
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).POS.Map, "?" & UserList(UserIndex).Char.CharIndex)
    End If
    Call UserDa�oNpc(UserIndex, NpcIndex)
Else
     Call SendData(ToPCArea, UserIndex, UserList(UserIndex).POS.Map, "-" & UserList(UserIndex).Char.CharIndex)
     Call SendData(ToIndex, UserIndex, 0, "U1")
End If

End Sub
Public Function TiempoTranscurrido(ByVal Desde As Single) As Single

TiempoTranscurrido = Timer - Desde

If TiempoTranscurrido < -5 Then
    TiempoTranscurrido = TiempoTranscurrido + 86400
ElseIf TiempoTranscurrido < 0 Then
    TiempoTranscurrido = 0
End If

End Function
Public Function UsuarioImpacto(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer) As Boolean
Dim ProbRechazo As Long
Dim Rechazo As Boolean
Dim ProbExito As Long
Dim PoderAtaque As Long
Dim UserPoderEvasion As Long
Dim proyectil As Boolean

If UserList(AtacanteIndex).Invent.WeaponEqpObjIndex = 0 Then
    proyectil = False
Else: proyectil = ObjData(UserList(AtacanteIndex).Invent.WeaponEqpObjIndex).proyectil = 1
End If

UserPoderEvasion = (1 + 0.05 * Buleano(UserList(VictimaIndex).Recompensas(3) = 2 And (UserList(VictimaIndex).Clase = ARQUERO Or UserList(VictimaIndex).Clase = NIGROMANTE))) _
                    * PoderEvasion(VictimaIndex)

If UserList(VictimaIndex).Invent.EscudoEqpObjIndex Then UserPoderEvasion = UserPoderEvasion + PoderEvasionEscudo(VictimaIndex)


If UserList(AtacanteIndex).Invent.WeaponEqpObjIndex Then
    If proyectil Then
        PoderAtaque = (1 + 0.05 * Buleano(UserList(AtacanteIndex).Clase = ARQUERO And UserList(AtacanteIndex).Recompensas(3) = 1) + 0.1 * Buleano(UserList(AtacanteIndex).Recompensas(3) = 1 And (UserList(AtacanteIndex).Clase = GUERRERO Or UserList(AtacanteIndex).Clase = CAZADOR))) _
        * PoderAtaqueProyectil(AtacanteIndex)
    Else
        PoderAtaque = (1 + 0.05 * Buleano(UserList(AtacanteIndex).Clase = PALADIN And UserList(AtacanteIndex).Recompensas(3) = 2)) _
        * PoderAtaqueArma(AtacanteIndex)
    End If
Else
    PoderAtaque = PoderAtaqueWresterling(UserList(AtacanteIndex).Stats.UserSkills(Wresterling), UserList(AtacanteIndex).Stats.UserAtributos(Agilidad), UserList(AtacanteIndex).Clase, UserList(AtacanteIndex).Stats.ELV)
End If

ProbExito = Maximo(10, Minimo(90, 50 + ((PoderAtaque - UserPoderEvasion) * 0.4)))

UsuarioImpacto = (RandomNumber(1, 100) <= ProbExito)


If UserList(VictimaIndex).Invent.EscudoEqpObjIndex Then
    
    
    If Not UsuarioImpacto Then
      ProbRechazo = Maximo(10, Minimo(90, 100 * (UserList(VictimaIndex).Stats.UserSkills(Defensa) / (UserList(VictimaIndex).Stats.UserSkills(Defensa) + UserList(VictimaIndex).Stats.UserSkills(Tacticas)))))
      Rechazo = (RandomNumber(1, 100) <= ProbRechazo)
      If Rechazo Then
            
            Call SendData(ToPCArea, AtacanteIndex, UserList(AtacanteIndex).POS.Map, "&" & UserList(AtacanteIndex).Char.CharIndex)
            Call SendData(ToIndex, AtacanteIndex, 0, "8")
            Call SendData(ToIndex, VictimaIndex, 0, "7")
            Call SubirSkill(VictimaIndex, Defensa, 25)
      End If
    End If
End If
    
If UsuarioImpacto Then
    If UserList(AtacanteIndex).Invent.WeaponEqpObjIndex Then
        If Not proyectil Then
            Call SubirSkill(AtacanteIndex, Armas)
        Else: Call SubirSkill(AtacanteIndex, Proyectiles)
        End If
    Else
        Call SubirSkill(AtacanteIndex, Wresterling)
    End If
End If

End Function
Public Sub UsuarioAtacaUsuario(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)

If Not PuedeAtacar(AtacanteIndex, VictimaIndex) Then Exit Sub

'If Distancia(UserList(AtacanteIndex).POS, UserList(VictimaIndex).POS) > MAXDISTANCIAARCO Then
'   Call SendData(ToIndex, AtacanteIndex, 0, "3G")
'   Exit Sub
'End If

Call UsuarioAtacadoPorUsuario(AtacanteIndex, VictimaIndex)

If UsuarioImpacto(AtacanteIndex, VictimaIndex) Then
    'If UserList(AtacanteIndex).flags.Invisible Then Call BajarInvisible(AtacanteIndex)
    
    Call SendData(ToPCArea, AtacanteIndex, UserList(AtacanteIndex).POS.Map, "TW" & "10")

    If UserList(VictimaIndex).flags.Navegando = 0 And Not UserList(VictimaIndex).flags.Meditando Then Call SendData(ToPCArea, VictimaIndex, UserList(VictimaIndex).POS.Map, "CFX" & UserList(VictimaIndex).Char.CharIndex & "," & FXSANGRE & "," & 0)
    
    Call UserDa�oUser(AtacanteIndex, VictimaIndex)
Else
    Call SendData(ToPCArea, AtacanteIndex, UserList(AtacanteIndex).POS.Map, "-" & UserList(AtacanteIndex).Char.CharIndex)
    Call SendData(ToIndex, AtacanteIndex, 0, "U1")
    Call SendData(ToIndex, VictimaIndex, 0, "U3" & UserList(AtacanteIndex).Name)
End If

End Sub
Public Sub UserDa�oUser(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)
Dim Da�o As Long, antda�o As Integer
Dim lugar As Integer, absorbido As Long
Dim defbarco As Integer
Dim Obj As ObjData
Dim Obj2 As ObjData
Dim j As Integer

Da�o = CalcularDa�o(AtacanteIndex)

antda�o = Da�o

If UserList(AtacanteIndex).flags.Navegando = 1 Then
     Obj = ObjData(UserList(AtacanteIndex).Invent.BarcoObjIndex)
     Da�o = Da�o + RandomNumber(Obj.MinHit, Obj.MaxHit)
End If

If UserList(VictimaIndex).flags.Navegando = 1 Then
     Obj = ObjData(UserList(VictimaIndex).Invent.BarcoObjIndex)
     defbarco = RandomNumber(Obj.MinDef, Obj.MaxDef)
End If

lugar = RandomNumber(1, 6)

Select Case lugar
  
  Case bCabeza
        If UserList(VictimaIndex).Invent.CascoEqpObjIndex Then
            If Not (UserList(AtacanteIndex).Clase = CAZADOR And UserList(AtacanteIndex).Recompensas(3) = 2) Then
                Obj = ObjData(UserList(VictimaIndex).Invent.CascoEqpObjIndex)
                absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
            End If
        End If
        
  Case Else
        If UserList(VictimaIndex).Invent.ArmourEqpObjIndex Then
           Obj = ObjData(UserList(VictimaIndex).Invent.ArmourEqpObjIndex)
           absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
        End If
        
        If UserList(VictimaIndex).Invent.EscudoEqpObjIndex Then
           Obj2 = ObjData(UserList(VictimaIndex).Invent.EscudoEqpObjIndex)
            absorbido = absorbido + RandomNumber(Obj2.MinDef, Obj2.MaxDef)
        End If
        
End Select

absorbido = absorbido + defbarco + 2 * Buleano(UserList(VictimaIndex).Clase = GUERRERO And UserList(VictimaIndex).Recompensas(2) = 2)
Da�o = Maximo(1, Da�o - absorbido)

Call SendData(ToIndex, AtacanteIndex, 0, "N5" & lugar & "," & Da�o & "," & UserList(VictimaIndex).Name)
Call SendData(ToIndex, VictimaIndex, 0, "N4" & lugar & "," & Da�o & "," & UserList(AtacanteIndex).Name)

If Not UserList(VictimaIndex).flags.Quest Then UserList(VictimaIndex).Stats.MinHP = UserList(VictimaIndex).Stats.MinHP - Da�o

If UserList(AtacanteIndex).flags.Hambre = 0 And UserList(AtacanteIndex).flags.Sed = 0 Then
    If UserList(AtacanteIndex).Invent.WeaponEqpObjIndex Then
        If ObjData(UserList(AtacanteIndex).Invent.WeaponEqpObjIndex).proyectil Then
            Call SubirSkill(AtacanteIndex, Proyectiles)
        Else: Call SubirSkill(AtacanteIndex, Armas)
        End If
    Else
        Call SubirSkill(AtacanteIndex, Wresterling)
    End If
    
    Call SubirSkill(AtacanteIndex, Tacticas)
    
    
    If PuedeApu�alar(AtacanteIndex) Then
        Call DoApu�alar(AtacanteIndex, 0, VictimaIndex, Da�o)
        Call SubirSkill(AtacanteIndex, Apu�alar)
    End If
End If

If UserList(VictimaIndex).Stats.MinHP <= 0 Then
     Call ContarMuerte(VictimaIndex, AtacanteIndex)
     
     

     For j = 1 To MAXMASCOTAS - 17 * Buleano(Not UserList(AtacanteIndex).flags.Quest)
        If UserList(AtacanteIndex).MascotasIndex(j) Then
            If Npclist(UserList(AtacanteIndex).MascotasIndex(j)).Target = VictimaIndex Then Npclist(UserList(AtacanteIndex).MascotasIndex(j)).Target = 0
            Call FollowAmo(UserList(AtacanteIndex).MascotasIndex(j))
        End If
     Next

     Call ActStats(VictimaIndex, AtacanteIndex)
End If
        


Call CheckUserLevel(AtacanteIndex)


End Sub
Sub UsuarioAtacadoPorUsuario(ByVal AttackerIndex As Integer, ByVal VictimIndex As Integer)
    If TriggerZonaPelea(AttackerIndex, VictimIndex) = TRIGGER7_PERMITE Then Exit Sub
If UserList(AttackerIndex).POS.Map = 35 Then Exit Sub
If UserList(AttackerIndex).POS.Map = 36 Then Exit Sub
If UserList(AttackerIndex).POS.Map = 86 Then Exit Sub

Call AllMascotasAtacanUser(AttackerIndex, VictimIndex)
Call AllMascotasAtacanUser(VictimIndex, AttackerIndex)

End Sub
Sub AllMascotasAtacanUser(ByVal victim As Integer, ByVal Maestro As Integer)
Dim iCount As Integer

For iCount = 1 To MAXMASCOTAS - 17 * Buleano(Not UserList(Maestro).flags.Quest)
    If UserList(Maestro).MascotasIndex(iCount) Then
        Npclist(UserList(Maestro).MascotasIndex(iCount)).flags.AttackedBy = victim
        Npclist(UserList(Maestro).MascotasIndex(iCount)).Movement = NPCDEFENSA
        Npclist(UserList(Maestro).MascotasIndex(iCount)).Hostile = 1
    End If
Next

End Sub
Public Function PuedeAtacar(ByVal AttackerIndex As Integer, ByVal VictimIndex As Integer) As Boolean

If UserList(AttackerIndex).flags.Muerto Then
    Call SendData(ToIndex, AttackerIndex, 0, "MU")
    Exit Function
End If

If UserList(VictimIndex).flags.Muerto Then
    Call SendData(ToIndex, AttackerIndex, 0, "0X")
    Exit Function
End If

If AttackerIndex = VictimIndex Then
    If UserList(AttackerIndex).flags.Privilegios = 3 Then
        PuedeAtacar = True
        Exit Function
    Else
        Call SendData(ToIndex, AttackerIndex, 0, "%3")
        Exit Function
    End If
End If

Dim T As Trigger7
T = TriggerZonaPelea(AttackerIndex, VictimIndex)

If UserList(VictimIndex).flags.Protegido = 1 Then
    Call SendData(ToIndex, AttackerIndex, 0, "||El usuario est� siendo protegido por un GM." & FONTTYPE_FIGHT)
    Exit Function
ElseIf UserList(VictimIndex).flags.Protegido = 2 Then
    Call SendData(ToIndex, AttackerIndex, 0, "||El usuario est� siendo protegido por conexi�n." & FONTTYPE_FIGHT)
    Exit Function
End If

If UserList(VictimIndex).flags.Privilegios >= 1 Then
    If UserList(VictimIndex).flags.AdminInvisible = 0 Then Call SendData(ToIndex, AttackerIndex, 0, "E0")
    Exit Function
End If

If UserList(AttackerIndex).flags.Protegido = 1 Then
    Call SendData(ToIndex, AttackerIndex, 0, "||No podes atacar mientr�s est�s siendo protegido por un GM." & FONTTYPE_FIGHT)
    Exit Function
ElseIf UserList(AttackerIndex).flags.Protegido = 2 Then
    Call SendData(ToIndex, AttackerIndex, 0, "||No podes atacar tan pronto al conectarte." & FONTTYPE_FIGHT)
    Exit Function
End If

If UserList(VictimIndex).POS.Map <> 190 Then
    If Not MapInfo(UserList(VictimIndex).POS.Map).Pk And Not TiempoTranscurrido(UserList(VictimIndex).Counters.LastRobo <= 10) Then
        Call SendData(ToIndex, AttackerIndex, 0, "7G")
        Exit Function
    End If
End If

If MapData(UserList(VictimIndex).POS.Map, UserList(VictimIndex).POS.X, UserList(VictimIndex).POS.Y).trigger = 4 Or MapData(UserList(AttackerIndex).POS.Map, UserList(AttackerIndex).POS.X, UserList(AttackerIndex).POS.Y).trigger = 4 Then
    Call SendData(ToIndex, AttackerIndex, 0, "8G")
    Exit Function
End If

If T = TRIGGER7_PERMITE Then
    PuedeAtacar = True
    Exit Function
ElseIf T = TRIGGER7_PROHIBE Then
    PuedeAtacar = False
    Exit Function
End If

If UserList(VictimIndex).POS.Map = 191 Then
    Call SendData(ToIndex, AttackerIndex, 0, "||No puedes atacar a otros usuarios en el mapa de espera de torneo." & FONTTYPE_FIGHT)
    Exit Function
End If

If UserList(VictimIndex).flags.Party And UserList(AttackerIndex).PartyIndex = UserList(VictimIndex).PartyIndex Then
    Call SendData(ToIndex, AttackerIndex, 0, "||No puedes atacar a miembros de tu party." & FONTTYPE_FIGHT)
    Exit Function
End If

If Not ModoQuest And Len(UserList(VictimIndex).GuildInfo.GuildName) > 0 And UserList(AttackerIndex).GuildInfo.GuildName = UserList(VictimIndex).GuildInfo.GuildName Then
    PuedeAtacar = True
    Exit Function
End If

If UserList(AttackerIndex).Faccion.Bando <> Neutral And UserList(AttackerIndex).Faccion.Bando = UserList(VictimIndex).Faccion.Bando Then
    If Len(UserList(AttackerIndex).GuildInfo.GuildName) = 0 Or Len(UserList(VictimIndex).GuildInfo.GuildName) = 0 Then
        Call SendData(ToIndex, AttackerIndex, 0, Mensajes(UserList(AttackerIndex).Faccion.Bando, 20))
        Exit Function
    End If
    If Not UserList(AttackerIndex).GuildRef.IsEnemy(UserList(VictimIndex).GuildInfo.GuildName) Then
        Call SendData(ToIndex, AttackerIndex, 0, Mensajes(UserList(AttackerIndex).Faccion.Bando, 20))
        Exit Function
    End If
    If ModoQuest Then
        Call SendData(ToIndex, AttackerIndex, 0, "||Durante una quest no puedes atacar a miembros de tu facci�n aunque pertenezcan a clanes enemigos." & FONTTYPE_FIGHT)
        Exit Function
    End If
End If

If EsNewbie(VictimIndex) And (EsNewbie(AttackerIndex) Or UserList(AttackerIndex).Faccion.Bando = Real) Then
    Call SendData(ToIndex, AttackerIndex, 0, "%1")
    Exit Function
End If

If EsNewbie(AttackerIndex) And UserList(VictimIndex).Faccion.Bando = Real Then
    Call SendData(ToIndex, AttackerIndex, 0, "%2")
    Exit Function
End If

PuedeAtacar = True

End Function
Public Function PuedeAtacarMascota(ByVal AttackerIndex As Integer, ByVal VictimIndex As Integer) As Boolean

If AttackerIndex = VictimIndex Then
    PuedeAtacarMascota = True
    Exit Function
End If

If UserList(VictimIndex).flags.Protegido = 1 Then
    Call SendData(ToIndex, AttackerIndex, 0, "||No puedes atacar mascotas de usuarios protegidos por GMs." & FONTTYPE_FIGHT)
    Exit Function
ElseIf UserList(VictimIndex).flags.Protegido = 2 Then
    Call SendData(ToIndex, AttackerIndex, 0, "||No puedes atacar mascotas de usuarios protegidos por conexi�n." & FONTTYPE_FIGHT)
    Exit Function
End If

If UserList(VictimIndex).flags.Privilegios >= 1 Then Exit Function

If UserList(AttackerIndex).flags.Protegido = 1 Then
    Call SendData(ToIndex, AttackerIndex, 0, "||No podes atacar mientr�s est�s siendo protegido por un GM." & FONTTYPE_FIGHT)
    Exit Function
ElseIf UserList(AttackerIndex).flags.Protegido = 2 Then
    Call SendData(ToIndex, AttackerIndex, 0, "||No podes atacar tan pronto al conectarte." & FONTTYPE_FIGHT)
    Exit Function
End If

If UserList(VictimIndex).POS.Map <> 190 Then
    If Not MapInfo(UserList(VictimIndex).POS.Map).Pk And Not TiempoTranscurrido(UserList(VictimIndex).Counters.LastRobo <= 10) Then
        Call SendData(ToIndex, AttackerIndex, 0, "||No podes atacar mascotas en zonas seguras." & FONTTYPE_FIGHT)
        Exit Function
    End If
End If

If MapData(UserList(VictimIndex).POS.Map, UserList(VictimIndex).POS.X, UserList(VictimIndex).POS.Y).trigger = 4 Or MapData(UserList(AttackerIndex).POS.Map, UserList(AttackerIndex).POS.X, UserList(AttackerIndex).POS.Y).trigger = 4 Then
    Call SendData(ToIndex, AttackerIndex, 0, "8G")
    Exit Function
End If

If Len(UserList(VictimIndex).GuildInfo.GuildName) > 0 And UserList(AttackerIndex).GuildInfo.GuildName = UserList(VictimIndex).GuildInfo.GuildName Then
    PuedeAtacarMascota = True
    Exit Function
End If

If UserList(AttackerIndex).Faccion.Bando <> Neutral And UserList(AttackerIndex).Faccion.Bando = UserList(VictimIndex).Faccion.Bando Then
    If Len(UserList(AttackerIndex).GuildInfo.GuildName) = 0 Or Len(UserList(VictimIndex).GuildInfo.GuildName) = 0 Then
        Call SendData(ToIndex, AttackerIndex, 0, "||No puedes atacar mascotas de tu bando." & FONTTYPE_INFO)
        Exit Function
    End If
    If Not UserList(AttackerIndex).GuildRef.IsEnemy(UserList(VictimIndex).GuildInfo.GuildName) Then
        Call SendData(ToIndex, AttackerIndex, 0, "||No puedes atacar mascotas de tu bando a menos que tu clan este en guerra con el del due�o." & FONTTYPE_INFO)
        Exit Function
    End If
End If

If EsNewbie(VictimIndex) And (EsNewbie(AttackerIndex) Or UserList(AttackerIndex).Faccion.Bando = Real) Then
    Call SendData(ToIndex, AttackerIndex, 0, "||Los miembros de la Alianza del F�nix no pueden atacar mascotas de newbies." & FONTTYPE_INFO)
    Exit Function
End If

If EsNewbie(AttackerIndex) And UserList(VictimIndex).Faccion.Bando = Real Then
    Call SendData(ToIndex, AttackerIndex, 0, "||Los newbies no pueden atacar mascotas de la Alianza del F�nix." & FONTTYPE_INFO)
    Exit Function
End If

If UserList(AttackerIndex).POS.Map = 190 Then Exit Function

PuedeAtacarMascota = True

End Function
Public Function PuedeRobar(ByVal AttackerIndex As Integer, ByVal VictimIndex As Integer) As Boolean

If UserList(AttackerIndex).flags.Muerto Then
    Call SendData(ToIndex, AttackerIndex, 0, "MU")
    Exit Function
End If

If UserList(VictimIndex).flags.Muerto Then
    Call SendData(ToIndex, AttackerIndex, 0, "0X")
    Exit Function
End If

If AttackerIndex = VictimIndex Then
    Call SendData(ToIndex, AttackerIndex, 0, "%3")
    Exit Function
End If

If UserList(VictimIndex).flags.Protegido = 1 Then
    Call SendData(ToIndex, AttackerIndex, 0, "||El usuario est� siendo protegido por un GM." & FONTTYPE_FIGHT)
    Exit Function
ElseIf UserList(VictimIndex).flags.Protegido = 2 Then
    Call SendData(ToIndex, AttackerIndex, 0, "||El usuario est� siendo protegido por conexi�n." & FONTTYPE_FIGHT)
    Exit Function
End If

If UserList(VictimIndex).flags.Privilegios >= 1 Then
    If UserList(VictimIndex).flags.AdminInvisible = 0 Then Call SendData(ToIndex, AttackerIndex, 0, "/F")
    Exit Function
End If

If UserList(AttackerIndex).flags.Protegido = 1 Then
    Call SendData(ToIndex, AttackerIndex, 0, "||No podes atacar mientr�s est�s siendo protegido por un GM." & FONTTYPE_FIGHT)
    Exit Function
ElseIf UserList(AttackerIndex).flags.Protegido = 2 Then
    Call SendData(ToIndex, AttackerIndex, 0, "||No podes atacar tan pronto al conectarte." & FONTTYPE_FIGHT)
    Exit Function
End If

If UserList(VictimIndex).POS.Map <> 190 Then
    If Not MapInfo(UserList(VictimIndex).POS.Map).Pk And Not TiempoTranscurrido(UserList(VictimIndex).Counters.LastRobo <= 10) Then
        Call SendData(ToIndex, AttackerIndex, 0, "/A")
        Exit Function
    End If
End If

If MapData(UserList(VictimIndex).POS.Map, UserList(VictimIndex).POS.X, UserList(VictimIndex).POS.Y).trigger = 4 Or MapData(UserList(AttackerIndex).POS.Map, UserList(AttackerIndex).POS.X, UserList(AttackerIndex).POS.Y).trigger = 4 Then
    Call SendData(ToIndex, AttackerIndex, 0, "/B")
    Exit Function
End If

If UserList(VictimIndex).POS.Map = 191 Then
    Call SendData(ToIndex, AttackerIndex, 0, "||No puedes atacar a otros usuarios en el mapa de espera de torneo." & FONTTYPE_FIGHT)
    Exit Function
End If

If UserList(VictimIndex).flags.Party And UserList(AttackerIndex).PartyIndex = UserList(VictimIndex).PartyIndex Then
    Call SendData(ToIndex, AttackerIndex, 0, "||No puedes atacar a miembros de tu party." & FONTTYPE_FIGHT)
    Exit Function
End If

If UserList(VictimIndex).Stats.MinSta < UserList(VictimIndex).Stats.MaxSta / 10 Then
    Call SendData(ToIndex, AttackerIndex, 0, "||No puedes atacar usuarios que tienen menos del 10% de su stamina total." & FONTTYPE_INFO)
    Exit Function
End If

If Len(UserList(VictimIndex).GuildInfo.GuildName) > 0 And UserList(AttackerIndex).GuildInfo.GuildName = UserList(VictimIndex).GuildInfo.GuildName Then
    PuedeRobar = True
    Exit Function
End If

If UserList(AttackerIndex).Faccion.Bando <> Neutral And UserList(AttackerIndex).Faccion.Bando = UserList(VictimIndex).Faccion.Bando Then
    If Len(UserList(AttackerIndex).GuildInfo.GuildName) = 0 Or Len(UserList(VictimIndex).GuildInfo.GuildName) = 0 Then
        Call SendData(ToIndex, AttackerIndex, 0, Mensajes(UserList(AttackerIndex).Faccion.Bando, 20))
        Exit Function
    End If
    If Not UserList(AttackerIndex).GuildRef.IsEnemy(UserList(VictimIndex).GuildInfo.GuildName) Then
        Call SendData(ToIndex, AttackerIndex, 0, Mensajes(UserList(AttackerIndex).Faccion.Bando, 20))
        Exit Function
    End If
End If

If EsNewbie(VictimIndex) And (EsNewbie(AttackerIndex) Or UserList(AttackerIndex).Faccion.Bando = Real) Then
    Call SendData(ToIndex, AttackerIndex, 0, "%1")
    Exit Function
End If

If EsNewbie(AttackerIndex) And UserList(VictimIndex).Faccion.Bando = Real Then
    Call SendData(ToIndex, AttackerIndex, 0, "%2")
    Exit Function
End If

PuedeRobar = True

Call AllMascotasAtacanUser(AttackerIndex, VictimIndex)
Call AllMascotasAtacanUser(VictimIndex, AttackerIndex)

End Function
Public Function TriggerZonaPelea(ByVal Origen As Integer, ByVal Destino As Integer) As Trigger7
 
If Origen > 0 And Destino > 0 And Origen <= UBound(UserList) And Destino <= UBound(UserList) Then
    If MapData(UserList(Origen).POS.Map, UserList(Origen).POS.X, UserList(Origen).POS.Y).trigger = 7 Or _
        MapData(UserList(Destino).POS.Map, UserList(Destino).POS.X, UserList(Destino).POS.Y).trigger = 7 Then
        If (MapData(UserList(Origen).POS.Map, UserList(Origen).POS.X, UserList(Origen).POS.Y).trigger = MapData(UserList(Destino).POS.Map, UserList(Destino).POS.X, UserList(Destino).POS.Y).trigger) Then
            TriggerZonaPelea = TRIGGER7_PERMITE
        Else
            TriggerZonaPelea = TRIGGER7_PROHIBE
        End If
    Else
        TriggerZonaPelea = TRIGGER7_AUSENTE
    End If
Else
    TriggerZonaPelea = TRIGGER7_AUSENTE
End If
 
End Function

