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

Private Server  As Network.Server
Private Time(2) As Single
Private Mapping() As Integer
Public DisconnectTimeout As Long

Public Sub Listen(ByVal Limit As Long, ByVal Address As String, ByVal Service As String)
    Set Server = New Network.Server
    ReDim Mapping(1 To MaxUsers) As Integer
    
    Call Server.Attach(AddressOf OnServerConnect, AddressOf OnServerClose, AddressOf OnServerSend, AddressOf OnServerRecv)
    
    Call Server.Listen(Limit, Address, Service)
End Sub

Public Sub Disconnect()
    Call Server.Close
End Sub

Public Sub Tick(ByVal Delta As Single)
    Time(0) = Time(0) + Delta
    Time(1) = Time(1) + Delta
    
    If (Time(0) >= TIME_RECV_FREQUENCY) Then
        Time(0) = 0
        
        Call Server.Poll
    End If
        
    If (Time(1) >= TIME_SEND_FREQUENCY) Then
        Time(1) = 0
        
        Call Server.Flush
    End If
End Sub

Public Sub Poll()
    Call Server.Poll
    Call Server.Flush
End Sub

Public Sub Send(ByVal UserIndex As Long, ByVal Buffer As Network.Writer)
    Call Server.Send(UserList(UserIndex).ConnID, False, Buffer)
End Sub

Public Sub Flush(ByVal UserIndex As Long)
    Call Server.Flush(UserList(UserIndex).ConnID)
End Sub

Private Sub OnServerConnect(ByVal Connection As Long, ByVal Address As String)
On Error GoTo OnServerConnect_Err:
  
    Debug.Print ("OnServerConnect connecting new user on id: " & Connection & " ip: " & Address)
    
    If Mapping(Connection) > 0 Then
        Debug.Print "Conflicto entre id de aurora y userindex existente. Connection = " & Connection & ", Mapping(Connection) = " & Mapping(Connection) & ". Proceda con precaucion."
    End If
    
    If Connection <= MaxUsers Then
    
        'Busca si esta banneada la ip
        For i = 1 To BanIps.Count
    
            If BanIps.Item(i) = Address Then
                Call WriteErrorMsg(NewIndex, "Su IP se encuentra bloqueada en este servidor.")
                Exit Sub
            End If
    
        Next i
        
        Dim FreeUser As Long
        FreeUser = NextOpenUser()
        If UserList(FreeUser).InUse Then
           Call LogError("Trying to use an user slot marked as in use! slot: " & FreeUser)
           FreeUser = NextOpenUser()
        End If
    
        UserList(FreeUser).ConnIDValida = True
        UserList(FreeUser).IP = Address
        UserList(FreeUser).ConnID = Connection
        UserList(FreeUser).Counters.OnConnectTimestamp = GetTickCount()
        
        Mapping(Connection) = FreeUser
        
        If FreeUser >= LastUser Then LastUser = FreeUser
        
        Call WriteConnected(Mapping(Connection).ArrayIndex)
    Else
        Call Kick(Connection, "El server se encuentra lleno en este momento. Disculpe las molestias ocasionadas.")
    End If
    
    Exit Sub
End Sub

Private Sub OnServerClose(ByVal Connection As Long)

    On Error GoTo OnServerClose_Err:
    
    Dim UserRef As t_UserReference

    UserRef = Mapping(Connection)

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

    Mapping(Connection) = 0

    Exit Sub
    
OnServerClose_Err:
    Call ForcedClose(UserRef.ArrayIndex, Connection)
    Call TraceError(Err.Number, Err.description, "modNetwork.OnServerClose", Erl)
End Sub

Private Sub OnServerSend(ByVal Connection As Long, ByVal Message As Network.Reader)
On Error GoTo OnServerSend_Err:
    
    Exit Sub
    
OnServerSend_Err:
    Call Kick(Connection)
    Debug.Print (Erl)
End Sub

Private Sub OnServerRecv(ByVal Connection As Long, ByVal Message As Network.Reader)
On Error GoTo OnServerRecv_Err:
    
    Dim UserRef As t_UserReference
    UserRef = Mapping(Connection)

    Call Protocol.HandleIncomingData(UserRef.ArrayIndex, Message)
    
    Exit Sub
    
OnServerRecv_Err:
    Call Kick(Connection)
    Call TraceError(Err.Number, Err.description, "modNetwork.OnServerRecv", Erl)
End Sub

Public Sub Kick(ByVal Connection As Long, Optional ByVal Message As String = vbNullString)
On Error GoTo Kick_ErrHandler:

    If (Message <> vbNullString) Then
        Dim UserRef As t_UserReference
        UserRef = Mapping(Connection)
        If UserRef > 0 Then
            Call Protocol_Writes.WriteErrorMsg(UserRef, Message)
            If UserList(UserRef).flags.UserLogged Then
                Call Cerrar_Usuario(UserRef)
            End If
        End If
    End If
        
    Call Server.Flush(Connection)
    Call Server.Kick(Connection, True)
    Exit Sub
