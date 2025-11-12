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

Private pItems As Collection   ' keyed Collection for values (key => value)
Private pKeys  As Collection   ' ordered list of keys (strings)

Private Sub Class_Initialize()
    Set pItems = New Collection
    Set pKeys = New Collection
End Sub

' Add: raise error if key exists (Dictionary-like Add)
Public Sub Add(key As String, value As Variant)
    If Me.Exists(key) Then
        err.Raise vbObjectError + 10000, "Map.Add", "Key already exists: " & key
    End If
    pItems.Add value, CStr(key)
    pKeys.Add CStr(key)
End Sub

' SetValue: set or replace value for key
Public Sub SetValue(key As String, value As Variant)
    If Me.Exists(key) Then
        Call Me.Remove(key)
    End If
    pItems.Add value, CStr(key)
    pKeys.Add CStr(key)
End Sub

' Exists
Public Function Exists(key As String) As Boolean
    On Error Resume Next
    Dim v As Variant
    vAssignment v, pItems(CStr(key))
    If err.Number = 0 Then
        Exists = True
    Else
        Exists = False
        err.Clear
    End If
    On Error GoTo 0
End Function

' GetValue (returns Empty if not exists)
Public Function GetValue(key As String) As Variant
    On Error Resume Next
    Dim v As Variant
    vAssignment v, pItems(CStr(key))
    If err.Number = 0 Then
        vAssignment GetValue, v
    Else
'        GetValue = Empty
        Set GetValue = Nothing
        err.Clear
    End If
    On Error GoTo 0
End Function

' Remove key (no error if not exists)
Public Sub Remove(key As String)
    On Error Resume Next
    pItems.Remove CStr(key)
    If err.Number = 0 Then
        ' remove key from pKeys
        Dim i As Long
        For i = 1 To pKeys.Count
            If CStr(pKeys(i)) = CStr(key) Then
                pKeys.Remove i
                Exit For
            End If
        Next i
    Else
        err.Clear
    End If
    On Error GoTo 0
End Sub

' Keys: return a Collection of keys (read-only copy)
Public Function keys() As Collection
    Dim out As New Collection
    Dim i As Long
    For i = 1 To pKeys.Count
        out.Add pKeys(i)
    Next i
    Set keys = out
End Function

' Items: return a Collection of values (in same order as keys)
Public Function items() As Collection
    Dim out As New Collection
    Dim i As Long
    For i = 1 To pKeys.Count
        out.Add pItems(CStr(pKeys(i)))
    Next i
    Set items = out
End Function

' Count property
Public Property Get Count() As Long
    Count = pKeys.Count
End Property

' Clear
Public Sub Clear()
    Set pItems = New Collection
    Set pKeys = New Collection
End Sub

' Clone (shallow clone of keys and values - values themselves are Variant references)
Public Function Clone() As Map
    Dim m As New Map
    Dim k As Variant
    For Each k In pKeys
        m.Add CStr(k), pItems(CStr(k))
    Next k
    Set Clone = m
End Function

Private Sub vAssignment(ByRef var As Variant, ByRef vValue As Variant)
    If IsObject(vValue) Then
        Set var = vValue
    Else
        var = vValue
    End If
End Sub
