VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ScopeStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Module: ScopeStack
Option Explicit

' ScopeStack
' - Maintains a collection of Map frames (pFrames).
' - Raw() returns clones (safe).
' - RawByRef() returns actual Map objects (by reference) to allow closure frames to share storage.
' - LoadRaw(frames) loads clones (normal).
' - LoadRawByRef(frames) loads by reference (used for closures).
' - SetValue now writes to the first existing frame (top->bottom) so captures are shared-write.
' - Push/Pop maintain lexical frames.

Private pFrames As Collection ' collection of Map objects (frames)

Private Sub Class_Initialize()
    Set pFrames = New Collection
    ' initial frame
    Dim m As New Map
    pFrames.Add m
End Sub

Public Function CaptureByReference() As ScopeStack
    ' Create a new ScopeStack whose frames reference the same Map objects
    ' currently stored in this ScopeStack's pFrames. Useful for closures.
    Dim s As New ScopeStack
    Dim refFrames As New Collection
    Dim i As Long
    For i = 1 To pFrames.Count
        refFrames.Add pFrames(i)   ' add the map object reference
    Next i
    s.LoadRawByRef refFrames
    Set CaptureByReference = s
End Function
' Return clones of frames (safe outward view)
Public Property Get Raw() As Collection
    Dim out As New Collection
    Dim i As Long
    For i = 1 To pFrames.Count
        out.Add pFrames(i).Clone
    Next i
    Set Raw = out
End Property

' Return actual Map references (use carefully: for closures only)
Public Property Get RawByRef() As Collection
    Dim out As New Collection
    Dim i As Long
    For i = 1 To pFrames.Count
        out.Add pFrames(i)
    Next i
    Set RawByRef = out
End Property

' Replace internal frames with clones of provided frames (default safe behavior)
Public Sub LoadRaw(framesCol As Collection) 'expects a collection
    Set pFrames = New Collection
    Dim f As Variant
    For Each f In framesCol
        If TypeName(f) = "Map" Then
            pFrames.Add f.Clone
        Else
            Dim tmp As New Map
            pFrames.Add tmp
        End If
    Next f
    If pFrames.Count = 0 Then
        Dim mm As New Map
        pFrames.Add mm
    End If
End Sub

' Replace internal frames by reference (no cloning). Use only when building closures.
Public Sub LoadRawByRef(framesCol As Collection)
    Set pFrames = New Collection
    Dim f As Variant
    For Each f In framesCol
        If TypeName(f) = "Map" Then
            pFrames.Add f
        Else
            Dim tmp As New Map
            pFrames.Add tmp
        End If
    Next f
    If pFrames.Count = 0 Then
        Dim mm2 As New Map
        pFrames.Add mm2
    End If
End Sub

Public Sub Push()
    Dim fr As New Map
    pFrames.Add fr
End Sub

Public Sub Pop()
    If pFrames.Count <= 1 Then Exit Sub
    pFrames.Remove pFrames.Count
End Sub

' Has: search top->bottom for key
Public Function Has(key As String) As Boolean
    Dim i As Long
    For i = pFrames.Count To 1 Step -1
        If pFrames(i).Exists(key) Then
            Has = True
            Exit Function
        End If
    Next i
    Has = False
End Function

' GetValue: search top->bottom
Public Function GetValue(key As String) As Variant
    Dim i As Long
    For i = pFrames.Count To 1 Step -1
        If pFrames(i).Exists(key) Then
            vAssignment GetValue, pFrames(i).GetValue(key)
            Exit Function
        End If
    Next i
    GetValue = Empty
End Function

' SetValue: write to the first frame where key exists (top->bottom). Otherwise put into top frame.
Public Sub SetValue(key As String, val As Variant)
    Dim i As Long
    For i = pFrames.Count To 1 Step -1
        If pFrames(i).Exists(key) Then
            pFrames(i).SetValue key, val
            Exit Sub
        
        End If
    Next i
    ' not found -> set in top frame
    Dim top As Map
    Set top = pFrames(pFrames.Count)
    top.SetValue key, val
End Sub

' Clone the ScopeStack (deep): returns a new ScopeStack with cloned frames
Public Function Clone() As ScopeStack
    Dim s As New ScopeStack
    Dim arr As New Collection
    Dim i As Long
    For i = 1 To pFrames.Count
        arr.Add pFrames(i).Clone
    Next i
    s.LoadRaw arr
    Set Clone = s
End Function

Private Sub vAssignment(ByRef var As Variant, ByRef vValue As Variant)
    If IsObject(vValue) Then
        Set var = vValue
    Else
        var = vValue
    End If
End Sub
