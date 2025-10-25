Attribute VB_Name = "mdlCanjes"
'################################################################################################
'################################### SALVAGUARDA DE PUNTOS ######################################
'################################################################################################
Public Sub GuardarPunto(Name As String, Cantidad As Integer)

Dim i As Integer
Dim VariableAire As Integer
Dim PuntoAire As Integer
Dim Agregardatos As String

For i = 1 To GetVar(App.Path & "\Sistemapuntos.log", "Beneficiarios", "Cantidad")
    If GetVar(App.Path & "\Sistemapuntos.log", "Ben" & i, "Nombre") = Name Then
        PuntoAire = GetVar(App.Path & "\Sistemapuntos.log", "Ben" & i, "Puntos") + Cantidad
        Agregardatos = str(PuntoAire)
        Call WriteVar(App.Path & "\Sistemapuntos.log", "Ben" & i, "Puntos", Agregardatos)
        Exit Sub
    End If
Next

VariableAire = GetVar(App.Path & "\Sistemapuntos.log", "Beneficiarios", "Cantidad") + 1
Agregardatos = str(VariableAire)
Call WriteVar(App.Path & "\Sistemapuntos.log", "Beneficiarios", "Cantidad", Agregardatos)
Agregardatos = str(PuntoAire)
Call WriteVar(App.Path & "\Sistemapuntos.log", "Ben" & VariableAire, "Nombre", Name)
Call WriteVar(App.Path & "\Sistemapuntos.log", "Ben" & VariableAire, "Puntos", Agregardatos)

End Sub

Public Sub EliminarPunto(Name As String, Cantidad As Integer)

Dim i As Integer
Dim PuntoAire As Integer
Dim Agregardatos As String

For i = 1 To GetVar(App.Path & "\Sistemapuntos.log", "Beneficiarios", "Cantidad")
    If GetVar(App.Path & "\Sistemapuntos.log", "Ben" & i, "Nombre") = Name Then
        PuntoAire = GetVar(App.Path & "\Sistemapuntos.log", "Ben" & i, "Puntos")
        If PuntoAire >= Cantidad Then
            PuntoAire = PuntoAire - Cantidad
            Agregardatos = str(PuntoAire)
            Call WriteVar(App.Path & "\Sistemapuntos.log", "Ben" & i, "Puntos", Agregardatos)
            Exit Sub
        End If
        Exit Sub
    End If
Next

End Sub
'################################################################################################
'##################################### SISTEMA DE CANJEO ########################################
'################################################################################################
Public Sub VerificarCanjeo(Nombre As String, Index As String)

Dim i As Integer

For i = 1 To GetVar(App.Path & "\Canjeos.log", "Items", "Cantidad")
    If Index = GetVar(App.Path & "\Canjeos.log", "Canje" & i, "Index") Then
        Call IniciarCanjeo(Nombre, GetVar(App.Path & "\Canjeos.log", "Canje" & i, "Index"), CInt(GetVar(App.Path & "\Canjeos.log", "Canje" & i, "Costo")), CInt(GetVar(App.Path & "\Canjeos.log", "Canje" & i, "Item")))
        Exit Sub
    End If
Next
'Mensaje que no se encontro en la lista del servidor el numero de Index
End Sub

Public Sub IniciarCanjeo(Name As String, Index As String, Costo As Integer, GhItem As Long)
Dim i As Integer
Dim Icanje As Obj

For i = 1 To GetVar(App.Path & "\Sistemapuntos.log", "Beneficiarios", "Cantidad")
    If Name = GetVar(App.Path & "\Sistemapuntos.log", "Ben" & i, "Nombre") Then
        If Costo <= GetVar(App.Path & "\Sistemapuntos.log", "Ben" & i, "Puntos") Then
            Icanje.Amount = 1 'Cantidad de Item
            Icanje.OBJIndex = GhItem 'Numero de Item
            Call MeterItemEnInventario(NameIndex(Name), Icanje)
            Call EliminarPunto(Name, Costo)
            Exit Sub
        End If
        'Mensaje no tiene suficientes puntos para canjear
        Exit Sub
    End If
Next

End Sub
