VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Map"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Module: Map
Option Explicit

Private Type Buffer
    Capacity As Long
    index As Long
    Storage() As Variant
End Type
Private pItems As Buffer   ' keyed Collection for values (key => value)
Private pKeys  As Buffer   ' ordered list of keys (strings)
Private emptyVariant As Variant
Private tmpBuffer As Variant

Private Sub Class_Initialize()
    BufferInit
End Sub

Private Sub BufferInit()
    pItems.Capacity = 64
    pItems.index = -1
    ReDim pItems.Storage(0 To pItems.Capacity - 1)
    pKeys.Capacity = 64
    pKeys.index = -1
    ReDim pKeys.Storage(0 To pKeys.Capacity - 1)
End Sub
Public Function Exists(key As String) As Boolean
    Dim i As Long
    Dim tmpResult As Boolean
    i = 0
    Do
        tmpResult = (pKeys.Storage(i) = key)
        i = i + 1
    Loop While i <= pKeys.index And Not tmpResult
    Exists = tmpResult
End Function

Private Function GetKeyIndex(key As String) As Long
    Dim i As Long
    Dim tmpBool As Boolean
    i = 0
    Do
        tmpBool = (pKeys.Storage(i) = key)
        i = i + 1
    Loop While i <= pKeys.index And Not tmpBool
    If tmpBool Then
        GetKeyIndex = i - 1
    Else
        GetKeyIndex = -1
    End If
End Function

Private Sub ExpandBuffer(ByRef aBuffer As Buffer)
    aBuffer.Capacity = 2 * (aBuffer.Capacity + 1)
    ReDim Preserve aBuffer.Storage(0 To aBuffer.Capacity - 1)
End Sub

' Add: raise error if key exists (Dictionary-like Add)
Public Sub Add(key As String, value As Variant)
    If Exists(key) Then
        err.Raise vbObjectError + 10000, "Map.Add", "Key already exists: " & key
    End If
    AppendBufferValue key, value
End Sub
Private Sub AppendBufferValue(key As String, value As Variant)
    pKeys.index = pKeys.index + 1
    pItems.index = pKeys.index
    On Error GoTo SetValue_errHandler
    pKeys.Storage(pKeys.index) = key
    vAssignment pItems.Storage(pKeys.index), value
    Exit Sub
SetValue_errHandler:
    err.Clear
    ExpandBuffer pKeys
    ExpandBuffer pItems
    pKeys.Storage(pKeys.index) = key
    pItems.Storage(pKeys.index) = value
End Sub

' SetValue: set or replace value for key
Public Sub SetValue(key As String, value As Variant)
    If Exists(key) Then 'Update
        Dim idx As Long
        idx = GetKeyIndex(key)
        vAssignment pItems.Storage(idx), value
    Else 'Append
        Add key, value
    End If
End Sub

' GetValue (returns Empty if not exists)
Public Property Get GetValue(key As String) As Variant
Attribute GetValue.VB_UserMemId = 0
    On Error Resume Next
    Dim idx As Long
    idx = GetKeyIndex(key)
    If idx > -1 Then
        vAssignment GetValue, pItems.Storage(idx)
    Else
        Set GetValue = Nothing
    End If
    On Error GoTo 0
End Property

Public Sub Remove(key As String)
    Dim iCounter As Long
    Dim pointer As Long
    Dim index As Long
    Dim curIdx As Long
    
    index = GetKeyIndex(key)
    curIdx = pKeys.index
    If index > -1 Then
        pKeys.Storage(index) = emptyVariant
        pItems.Storage(index) = emptyVariant
        For iCounter = index + 1 To curIdx
            Swap pKeys, index + pointer, iCounter
            Swap pItems, index + pointer, iCounter
            pointer = pointer + 1
        Next
        pKeys.index = pKeys.index - 1
        pItems.index = pKeys.index
    End If
End Sub
Private Sub Swap(ByRef aBuffer As Buffer, position1 As Long, position2 As Long)
    If aBuffer.index >= 0 Then
        vAssignment tmpBuffer, aBuffer.Storage(position1)
        vAssignment aBuffer.Storage(position1), aBuffer.Storage(position2)
        vAssignment aBuffer.Storage(position2), tmpBuffer
    End If
End Sub
' Keys: return a key given an index (read-only copy)
Public Function key(index As Long) As String
    key = pKeys.Storage(index - 1)
End Function
' Keys: return a Collection of keys (read-only copy)
Public Function keys() As Collection
    Dim out As New Collection
    Dim i As Long
    For i = 0 To pKeys.index
        out.Add pKeys.Storage(i)
    Next i
    Set keys = out
End Function
' Items: return a Collection of values (in same order as keys)
Public Function items() As Collection
    Dim out As New Collection
    Dim i As Long
    For i = 0 To pKeys.index
        out.Add pItems.Storage(i)
    Next i
    Set items = out
End Function
' Count property
Public Property Get count() As Long
    count = pKeys.index + 1
End Property
' Clear
Public Sub Clear()
    BufferInit
End Sub
' Clone (shallow clone of keys and values - values themselves are Variant references)
Public Function Clone() As Map
    Dim M As New Map
    Dim k As Long
    For k = 0 To pKeys.index
        M.Add CStr(pKeys.Storage(k)), pItems.Storage(k)
    Next k
    Set Clone = M
End Function
Private Sub vAssignment(ByRef var As Variant, ByRef vValue As Variant)
    If IsObject(vValue) Then
        Set var = vValue
    Else
        var = vValue
    End If
End Sub



