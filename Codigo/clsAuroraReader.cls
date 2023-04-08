VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAuroraReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************
' clsAuroraReader.cls
' Esta clase reemplaza a la anterior para convertirse en una especie de wrapper para el Reader de aurora
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

Private data                        As Network.Reader

Private queueCapacity               As Long

Private queueLength                 As Long

Private Sub Class_Initialize()
    '***************************************************
    'Autor: Walter Cristian Gonzalez (Kojax)
    'Ultima modificación: 06/04/2023
    '***************************************************
    queueCapacity = 10240
End Sub

''
' Clean up and release resources

Private Sub Class_Terminate()
    Set data = Nothing
End Sub

Public Sub CopyBuffer(ByRef source As clsAuroraReader)

    Set data = source.Reader
    queueLength = data.GetAvailable
    queueCapacity = data.GetCapacity
End Sub

Public Sub NewReader(ByRef buffer As Network.Reader)
    Set data = buffer
    queueLength = data.GetAvailable
    queueCapacity = data.GetCapacity
End Sub

Public Function ReadByte() As Byte
    ReadByte = data.ReadInt8
    queueLength = queueLength - 1
End Function

Public Function ReadInteger() As Integer
    ReadInteger = data.ReadInt16
    queueLength = queueLength - 2
End Function

Public Function ReadLong() As Long
    ReadLong = data.ReadInt32
    queueLength = queueLength - 4
End Function

Public Function ReadSingle() As Single
    ReadSingle = data.ReadReal32
    queueLength = queueLength - 4
End Function

Public Function ReadDouble() As Double
    ReadDouble = data.ReadReal64
    queueLength = queueLength - 8
End Function

Public Function ReadBoolean() As Boolean
    ReadBoolean = data.ReadBool
    queueLength = queueLength - 1
End Function

Public Function ReadASCIIString() As String
    Dim text As String
    text = data.ReadString8
    ReadASCIIString = data.ReadString8
    queueLength = queueLength - LenB(text) - 2
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
        Call data.Skip(Length)
        queueLength = 0
        ReadASCIIStringFixed = StrConv(buf, vbUnicode)
    Else
        Call Err.Raise(NOT_ENOUGH_DATA)

    End If

End Function

Public Function PeekByte() As Byte
    PeekByte = data.ReadInt8
    data.Skip -1
End Function

Public Function PeekInteger() As Integer
    PeekInteger = data.ReadInt16
    data.Skip -2
End Function

Public Function PeekLong() As Long
    PeekLong = data.ReadInt32
    data.Skip -4
End Function

Public Function PeekSingle() As Single
    PeekSingle = data.ReadReal32
    data.Skip -4
End Function

Public Function PeekDouble() As Double
    PeekDouble = data.ReadReal64
    data.Skip -8
End Function

Public Function PeekBoolean() As Boolean
    PeekBoolean = data.ReadBool
    data.Skip -1
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

Public Function Reader() As Variant
    Reader = data
End Function
