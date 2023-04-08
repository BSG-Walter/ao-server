Attribute VB_Name = "modNetwork"
'********************* COPYRIGHT NOTICE*********************
' Copyright (c) 2021-22 Martin Trionfetti, Pablo Marquez
' www.ao20.com.ar
' All rights reserved.
' Refer to licence for conditions of use.
' This copyright notice must always be left intact.
'****************** END OF COPYRIGHT NOTICE*****************
'
Option Explicit

Private Const TIME_RECV_FREQUENCY As Long = 0  ' In milliseconds
Private Const TIME_SEND_FREQUENCY As Long = 0 ' In milliseconds

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Server  As Network.Server
Private TimeTick(2) As Single
Private Mapping() As Integer
Public DisconnectTimeout As Long

Public Sub Listen(ByVal Limit As Long, ByVal Address As String, ByVal Service As String)
    Set Server = New Network.Server
    ReDim Mapping(1 To MaxUsers) As Integer
    
    Call Server.Attach(AddressOf OnServerConnect, AddressOf OnServerClose, AddressOf OnServerSend, AddressOf OnServerRecv)
    
    Call Server.Listen(Limit, Address, Service)
    frmMain.txtStatus.Text = Date & " " & Time & " - Escuchando conexiones entrantes ..."
    Debug.Print ("iniciado")
End Sub

Public Sub Disconnect()
    Call Server.Close
End Sub

Public Sub Connect()
    Call Listen(MaxUsers, "0.0.0.0", CStr(Puerto))
End Sub

Public Sub Tick(ByVal Delta As Single)
    TimeTick(0) = TimeTick(0) + Delta
    TimeTick(1) = TimeTick(1) + Delta
    
    If (Time(0) >= TIME_RECV_FREQUENCY) Then
        TimeTick(0) = 0
        
        Call Server.Poll
    End If
        
    If (Time(1) >= TIME_SEND_FREQUENCY) Then
        TimeTick(1) = 0
        
        Call Server.Flush
    End If
End Sub

Public Sub Poll()
    Call Server.Poll
    Call Server.Flush
End Sub

Public Sub Send(ByVal UserIndex As Long, ByVal buffer As Network.Writer)
    Call Server.Send(UserList(UserIndex).ConnID, False, buffer)
End Sub

Public Sub Flush(ByVal UserIndex As Long)
    Call Server.Flush(UserList(UserIndex).ConnID)
End Sub

Private Sub OnServerConnect(ByVal connection As Long, ByVal Address As String)
On Error GoTo OnServerConnect_Err:
  
    Debug.Print ("OnServerConnect connecting new user on id: " & connection & " ip: " & Address)
    
    If Mapping(connection) > 0 Then
        Debug.Print "Conflicto entre id de aurora y userindex existente. Connection = " & connection & ", Mapping(Connection) = " & Mapping(connection) & ". Proceda con precaucion."
    End If
    
    If connection <= MaxUsers Then
        Dim FreeUser As Long
        FreeUser = NextOpenUser()
        
        Dim i As Integer
        'Busca si esta banneada la ip
        For i = 1 To BanIps.Count
    
            If BanIps.Item(i) = Address Then
                Call WriteErrorMsg(FreeUser, "Su IP se encuentra bloqueada en este servidor.")
                Exit Sub
            End If
    
        Next i
    
        UserList(FreeUser).ConnIDValida = True
        UserList(FreeUser).IP = Address
        UserList(FreeUser).ConnID = connection
        
        Mapping(connection) = FreeUser
        
        If FreeUser >= LastUser Then LastUser = FreeUser
        
        Call WriteErrorMsg(FreeUser, "Probando")
        Call Send(FreeUser, UserList(FreeUser).outgoingData.Writer)
    Else
        Call Kick(connection, "El server se encuentra lleno en este momento. Disculpe las molestias ocasionadas.")
    End If
    
    Exit Sub
OnServerConnect_Err:
    Call Kick(connection)
End Sub

Private Sub OnServerClose(ByVal connection As Long)

    On Error GoTo OnServerClose_Err:
    
    Dim UserRef As Integer

    UserRef = Mapping(connection)

    If UserRef > 0 Then
        If UserList(UserRef).flags.UserLogged Then
            Call CloseSocketSL(UserRef)
            Call Cerrar_Usuario(UserRef)
        Else
            Call CloseSocket(UserRef)
        End If
    
        UserList(UserRef).ConnIDValida = False
        UserList(UserRef).ConnID = 0
    End If

    Mapping(connection) = 0

    Exit Sub
    
OnServerClose_Err:
    Call CloseSocket(UserRef)
    Debug.Print (Err.description & " modNetwork.OnServerClose")
End Sub

Private Sub OnServerSend(ByVal connection As Long, ByVal Message As Network.Reader)
On Error GoTo OnServerSend_Err:
    
    Exit Sub
    
OnServerSend_Err:
    Call Kick(connection)
    Debug.Print (Erl)
End Sub

Private Sub OnServerRecv(ByVal connection As Long, ByVal Message As Network.Reader)
On Error GoTo OnServerRecv_Err:
    Dim Slot As Integer
    Slot = Mapping(connection)
    
    #If AntiExternos Then
        'Notas de kojax:
        'aprovecho la funcion de getdata y setdata de aurora para seguir usando el cifrado XOR
        Dim Datos() As Byte
        Call Message.GetData(Datos)
        
        If UserList(Slot).flags.UserLogged Then
            Security.NAC_D_Byte Datos, UserList(Slot).Redundance
        Else
            Security.NAC_D_Byte Datos, 13
        End If
        
        Call Message.SetData(Datos)

    #End If
    

    Call Protocol.HandleIncomingData(Slot, Message)
    
    Exit Sub
    
OnServerRecv_Err:
    Call Kick(connection)
    Debug.Print (Err.description & "modNetwork.OnServerRecv")
End Sub

Public Sub Kick(ByVal connection As Long, Optional ByVal Message As String = vbNullString)
On Error GoTo Kick_ErrHandler:

    If (Message <> vbNullString) Then
        Dim UserRef As Integer
        UserRef = Mapping(connection)
        If UserRef > 0 Then
            Call WriteErrorMsg(UserRef, Message)
            Call Send(UserRef, UserList(UserRef).outgoingData.Writer)
            If UserList(UserRef).flags.UserLogged Then
                Call Cerrar_Usuario(UserRef)
            End If
        End If
    End If
        
    Call Server.Flush(connection)
    Call Server.Kick(connection, True)
    Exit Sub
Kick_ErrHandler:
    Debug.Print (Err.description & " modNetwork.Kick")
End Sub

' Test the time since last call and update the time
' log if there time betwen calls exced the limit
Public Sub PerformTimeLimitCheck(ByRef timer As Long, ByRef TestText As String, Optional ByVal TimeLimit As Long = 1000)
    Dim CurrTime As Long
    CurrTime = GetTickCount()
    If CurrTime - timer > TimeLimit Then
        Debug.Print ("Performance warning at: " & TestText & " elapsed time: " & CurrTime - timer)
    End If
    timer = GetTickCount()
End Sub

Public Sub ReiniciarSockets()

    Dim i As Long

    'Cierra el socket de escucha
    Call Disconnect
    
    'Cierra todas las conexiones
    For i = 1 To MaxUsers

        If UserList(i).ConnID <> -1 And UserList(i).ConnIDValida Then
            Call CloseSocket(i)
        End If

    Next i
    
    For i = 1 To MaxUsers
        Set UserList(i).incomingData = Nothing
        Set UserList(i).outgoingData = Nothing
    Next i
    
    ' No 'ta el PRESERVE :p
    ReDim UserList(1 To MaxUsers)

    For i = 1 To MaxUsers
        UserList(i).ConnID = -1
        UserList(i).ConnIDValida = False
        
        Set UserList(i).incomingData = New clsAuroraReader
        Set UserList(i).outgoingData = New clsAuroraWriter
    Next i
    
    LastUser = 1
    NumUsers = 0
    
    Call Sleep(100)
    
    Call Listen(MaxUsers, "0.0.0.0", CStr(Puerto))

End Sub

Public Sub LogNetworkSock(ByVal Str As String)

    On Error GoTo errHandler

    Dim nfile As Integer
        nfile = FreeFile ' obtenemos un canal
        
    Open App.Path & "\logs\network.log" For Append Shared As #nfile
        Print #nfile, Date & " " & Time & " " & Str
    Close #nfile

    Exit Sub

errHandler:

End Sub

