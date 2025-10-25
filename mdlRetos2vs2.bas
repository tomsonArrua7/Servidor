Attribute VB_Name = "mdlRetos2vs2"
Sub VerificarRetos(ByVal UserIndex As Integer)
If UserList(UserIndex).Retano.Retando_2 Then
    UserList(OPCDuelos.J1).Retano.Received_Request = False
    UserList(OPCDuelos.J1).Retano.Retando_2 = False
    UserList(OPCDuelos.J1).Retano.Send_Request = False
   
    UserList(OPCDuelos.J2).Retano.Received_Request = False
    UserList(OPCDuelos.J2).Retano.Retando_2 = False
    UserList(OPCDuelos.J2).Retano.Send_Request = False
   
    UserList(OPCDuelos.J3).Retano.Received_Request = False
    UserList(OPCDuelos.J3).Retano.Retando_2 = False
    UserList(OPCDuelos.J3).Retano.Send_Request = False
   
    UserList(OPCDuelos.J4).Retano.Received_Request = False
    UserList(OPCDuelos.J4).Retano.Retando_2 = False
    UserList(OPCDuelos.J4).Retano.Send_Request = False
   
    Call WarpUserChar(OPCDuelos.J1, ULLATHORPE.Map, ULLATHORPE.X + 1, ULLATHORPE.Y + 1, True)
    Call WarpUserChar(OPCDuelos.J2, ULLATHORPE.Map, ULLATHORPE.X + 1, ULLATHORPE.Y - 1, True)
    Call WarpUserChar(OPCDuelos.J3, ULLATHORPE.Map, ULLATHORPE.X - 1, ULLATHORPE.Y + 1, True)
    Call WarpUserChar(OPCDuelos.J4, ULLATHORPE.Map, ULLATHORPE.X - 1, ULLATHORPE.Y - 1, True)
 
    Call SendData(ToAll, 0, 0, "||2vs2: El reto se cancela porque " & UserList(UserIndex).Name & " desconectó." & FONTTYPE_TALK)
 
    frmMain.TimerRetos2v2.Enabled = False '> CUANDO CREEN EL TIMER, CAMBIENLEN EL NOMBRE.
    OPCDuelos.OCUP = False
    OPCDuelos.Tiempo = 0
    OPCDuelos.J1 = 0
    OPCDuelos.J2 = 0
    OPCDuelos.J3 = 0
    OPCDuelos.J4 = 0
End If
End Sub

Sub MuereReto2v2(ByVal UserIndex As Integer)

If UserList(UserIndex).Retano.Retando_2 = True Then
    If UserList(OPCDuelos.J1).flags.Muerto And UserList(OPCDuelos.J2).flags.Muerto Then 'Pareja 1 = muerta.
        Call SendData(ToAll, 0, 0, "||2vs2: " & UserList(OPCDuelos.J3).Name & " y " & UserList(OPCDuelos.J4).Name & _
                " derrotan a " & UserList(OPCDuelos.J1).Name & " y " & UserList(OPCDuelos.J2).Name & FONTTYPE_TALK)
              UserList(OPCDuelos.J3).flags.Reto = UserList(OPCDuelos.J3).flags.Reto + 1
        UserList(OPCDuelos.J4).flags.Reto = UserList(OPCDuelos.J4).flags.Reto + 1
        'Reseteamos flags.
        UserList(OPCDuelos.J1).Retano.Retando_2 = False
        UserList(OPCDuelos.J1).Retano.Send_Request = False
        UserList(OPCDuelos.J1).Retano.Received_Request = False
       
        UserList(OPCDuelos.J2).Retano.Retando_2 = False
        UserList(OPCDuelos.J2).Retano.Send_Request = False
        UserList(OPCDuelos.J2).Retano.Received_Request = False
 
        UserList(OPCDuelos.J3).Retano.Send_Request = False
        UserList(OPCDuelos.J3).Retano.Received_Request = False
        UserList(OPCDuelos.J3).Retano.Retando_2 = False
 
        UserList(OPCDuelos.J4).Retano.Send_Request = False
        UserList(OPCDuelos.J4).Retano.Received_Request = False
        UserList(OPCDuelos.J4).Retano.Retando_2 = False
      
        Call WarpUserChar(OPCDuelos.J1, ULLATHORPE.Map, ULLATHORPE.X, ULLATHORPE.Y) 'los mando a ulla
        Call WarpUserChar(OPCDuelos.J2, ULLATHORPE.Map, ULLATHORPE.X + 1, ULLATHORPE.Y) 'los mando a ulla
        Call WarpUserChar(OPCDuelos.J3, ULLATHORPE.Map, ULLATHORPE.X - 1, ULLATHORPE.Y) 'los mando a ulla
        Call WarpUserChar(OPCDuelos.J4, ULLATHORPE.Map, ULLATHORPE.X, ULLATHORPE.Y + 1) 'los mando a ulla
       
       OPCDuelos.ParejaEspera = False
       OPCDuelos.OCUP = False
 
       OPCDuelos.J1 = 0
       OPCDuelos.J2 = 0
       OPCDuelos.J3 = 0
       OPCDuelos.J4 = 0
       

        
    ElseIf UserList(OPCDuelos.J3).flags.Muerto And UserList(OPCDuelos.J4).flags.Muerto Then
        Call SendData(ToAll, 0, 0, "||2vs2: " & UserList(OPCDuelos.J1).Name & " y " & UserList(OPCDuelos.J2).Name & _
                " derrotan a " & UserList(OPCDuelos.J3).Name & " y " & UserList(OPCDuelos.J4).Name & FONTTYPE_TALK)
                UserList(OPCDuelos.J1).flags.Reto = UserList(OPCDuelos.J1).flags.Reto + 1
        UserList(OPCDuelos.J2).flags.Reto = UserList(OPCDuelos.J2).flags.Reto + 1
        'Reseteamos flags.
        UserList(OPCDuelos.J1).Retano.Retando_2 = False
        UserList(OPCDuelos.J1).Retano.Send_Request = False
        UserList(OPCDuelos.J1).Retano.Received_Request = False
       
        UserList(OPCDuelos.J2).Retano.Retando_2 = False
        UserList(OPCDuelos.J2).Retano.Send_Request = False
        UserList(OPCDuelos.J2).Retano.Received_Request = False
 
        UserList(OPCDuelos.J3).Retano.Send_Request = False
        UserList(OPCDuelos.J3).Retano.Received_Request = False
        UserList(OPCDuelos.J3).Retano.Retando_2 = False
 
        UserList(OPCDuelos.J4).Retano.Send_Request = False
        UserList(OPCDuelos.J4).Retano.Received_Request = False
        UserList(OPCDuelos.J4).Retano.Retando_2 = False
       
        Call WarpUserChar(OPCDuelos.J1, ULLATHORPE.Map, ULLATHORPE.X, ULLATHORPE.Y) 'los mando a ulla
        Call WarpUserChar(OPCDuelos.J2, ULLATHORPE.Map, ULLATHORPE.X + 1, ULLATHORPE.Y) 'los mando a ulla
        Call WarpUserChar(OPCDuelos.J3, ULLATHORPE.Map, ULLATHORPE.X - 1, ULLATHORPE.Y) 'los mando a ulla
        Call WarpUserChar(OPCDuelos.J4, ULLATHORPE.Map, ULLATHORPE.X, ULLATHORPE.Y + 1) 'los mando a ulla
       
        OPCDuelos.ParejaEspera = False
        OPCDuelos.OCUP = False
        OPCDuelos.J1 = 0
        OPCDuelos.J2 = 0
        OPCDuelos.J3 = 0
        OPCDuelos.J4 = 0
       
       
    End If
 
End If
 End Sub

