VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAuroraWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************
' clsAuroraWriter.cls
' Esta clase reemplaza a la anterior para convertirse en una especie de wrapper para el Writer de aurora
' asi nos ahorramos tener que refactorizar gran parte del protocolo binario
'
' Autor de la clase original: Juan Martin Sotuyo Dodero (Maraxus)
' Autor de este re-write: Walter Cristian Gonzalez (Kojax)
' (https://github.com/BSG-Walter)
'**************************************************************

Option Explicit
Option Base 0

' The error number thrown when there is not enough data in
' the buffer to read the specified data type.
' It's 9 (subscript out of range) + the object error constant
Private Const NOT_ENOUGH_DATA  As Long = vbObjectError + 9

' The error number thrown when there is not enough space in
' the buffer to write.
Private Const NOT_ENOUGH_SPACE As Long = vbObjectError + 10

Private Declare Sub CopyMemory _
                Lib "kernel32" _
                Alias "RtlMoveMemory" (ByRef destination As Any, _
                                       ByRef source As Any, _
                                       ByVal Length As Long)

Private data                        As Network.Writer

Private queueCapacity               As Long

Private queueLength                 As Long

Private Sub Class_Initialize()
    '***************************************************
    'Autor: Walter Cristian Gonzalez (Kojax)
    'Ultima modificación: 06/04/2023
    '***************************************************
    Set data = New Network.Writer
    queueLength = 0
    queueCapacity = data.GetCapacity
End Sub

''
' Clean up and release resources

Private Sub Class_Terminate()
    Set data = Nothing
End Sub

Public Sub CopyBuffer(ByRef source As clsAuroraWriter)

    Set data = source.Writer
    queueLength = data.GetAvailable
    queueCapacity = data.GetCapacity
End Sub

Public Function WriteByte(ByVal Value As Byte) As Long

    If queueCapacity - queueLength - 1 < 0 Then
        Call Err.Raise(NOT_ENOUGH_SPACE)
        Exit Function

    End If
    
    Call data.WriteInt8(Value)
    queueLength = queueLength + 1
    WriteByte = 1
End Function

Public Function WriteInteger(ByVal Value As Integer) As Long

    If queueCapacity - queueLength - 2 < 0 Then
        Call Err.Raise(NOT_ENOUGH_SPACE)
        Exit Function

    End If
    
    Call data.WriteInt16(Value)
    queueLength = queueLength + 2
    WriteInteger = 2
End Function

Public Function WriteLong(ByVal Value As Long) As Long

    If queueCapacity - queueLength - 4 < 0 Then
        Call Err.Raise(NOT_ENOUGH_SPACE)
        Exit Function

    End If
    
    Call data.WriteInt32(Value)
    queueLength = queueLength + 4
    WriteLong = 4
End Function

Public Function WriteSingle(ByVal Value As Single) As Long

    If queueCapacity - queueLength - 4 < 0 Then
        Call Err.Raise(NOT_ENOUGH_SPACE)
        Exit Function

    End If
    
    Call data.WriteReal32(Value)
    queueLength = queueLength + 4
    WriteSingle = 4
End Function


Public Function WriteDouble(ByVal Value As Double) As Long

    If queueCapacity - queueLength - 8 < 0 Then
        Call Err.Raise(NOT_ENOUGH_SPACE)
        Exit Function

    End If
    
    Call data.WriteReal64(Value)
    queueLength = queueLength + 8
    WriteDouble = 8
End Function

Public Function WriteBoolean(ByVal Value As Boolean) As Long

    If queueCapacity - queueLength - 1 < 0 Then
        Call Err.Raise(NOT_ENOUGH_SPACE)
        Exit Function

    End If
    
    Call data.WriteBool(Value)
    queueLength = queueLength + 1
    WriteBoolean = 1
End Function

Public Function WriteASCIIStringFixed(ByVal Value As String) As Long

    '***************************************************
    'Autor: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 04/27/06
    'Writes a fixed length ASCII string to the queue
    '***************************************************
    
    If queueCapacity - queueLength - Len(Value) < 0 Then
        Call Err.Raise(NOT_ENOUGH_SPACE)
        Exit Function

    End If
    
    Dim buf() As Byte

    ReDim buf(Len(Value) - 1) As Byte
    
    'Copy data to temp buffer
    Call CopyMemory(buf(0), ByVal StrPtr(StrConv(Value, vbFromUnicode)), Len(Value))
    
    Call data.WriteSafeArrayInt8(buf)
    queueLength = queueLength + Len(Value)
    WriteASCIIStringFixed = Len(Value)
End Function

Public Function WriteASCIIString(ByVal Value As String) As Long

    If queueCapacity - queueLength - Len(Value) - 2 < 0 Then
        Call Err.Raise(NOT_ENOUGH_SPACE)
        Exit Function

    End If
    
    Call data.WriteString8(Value)
    queueLength = queueLength + Len(Value) + 2
    WriteASCIIString = Len(Value) + 2
End Function

Public Function ReadASCIIStringFixed(ByVal Length As Long) As String

    '***************************************************
    'Autor: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 04/27/06
    'Reads a fixed length ASCII string from the queue and removes it
    '***************************************************
    If Length <= 0 Then Exit Function
    
    If queueLength >= Length Then

        Dim buf() As Byte

        ReDim buf(Length - 1) As Byte
        
        Call data.GetData(buf)
        queueLength = 0
        Call data.Clear
        ReadASCIIStringFixed = StrConv(buf, vbUnicode)
    Else
        Call Err.Raise(NOT_ENOUGH_DATA)

    End If

End Function

''
' Retrieves the current capacity of the queue.
'
' @return   The current capacity of the queue.

Public Property Get Capacity() As Long
    '***************************************************
    'Autor: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 04/27/06
    'Retrieves the current capacity of the queue
    '***************************************************
    Capacity = queueCapacity

End Property

''
' Sets the capacity of the queue.
'
' @param value The new capacity of the queue.
' @remarks If the new capacity is smaller than the current Length, all exceeding data is lost.
' @see Length

Public Property Let Capacity(ByVal Value As Long)
    '***************************************************
    'Autor: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 04/27/06
    'Sets the current capacity of the queue.
    'All data in the queue exceeding the new capacity is lost
    '***************************************************
    'Upate capacity
    queueCapacity = Value
    
    'All extra data is lost
    If Length > Value Then queueLength = Value

End Property

''
' Retrieves the length of the total data in the queue.
'
' @return   The length of the total data in the queue.

Public Property Get Length() As Long
    '***************************************************
    'Autor: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 04/27/06
    'Retrieves the current number of bytes in the queue
    '***************************************************
    Length = queueLength

End Property

''
' Retrieves the NOT_ENOUGH_DATA error code.
'
' @return   NOT_ENOUGH_DATA.

Public Property Get NotEnoughDataErrCode() As Long
    '***************************************************
    'Autor: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 04/27/06
    'Retrieves the NOT_ENOUGH_DATA error code
    '***************************************************
    NotEnoughDataErrCode = NOT_ENOUGH_DATA

End Property

Public Property Get NotEnoughSpaceErrCode() As Long
    '***************************************************
    'Autor: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 04/27/06
    'Retrieves the NOT_ENOUGH_SPACE error code
    '***************************************************
    NotEnoughSpaceErrCode = NOT_ENOUGH_SPACE

End Property

Public Function Writer(Optional ByVal dummy As Variant) As Variant
    Writer = data
End Function

Public Sub Clear()
    data.Clear
    queueLength = 0
    queueCapacity = data.GetCapacity
End Sub
