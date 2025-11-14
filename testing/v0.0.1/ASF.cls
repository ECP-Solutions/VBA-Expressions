VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ASF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private GLOBALS_ As Globals
Private Compiler_ As Compiler
Private VM_ As VM
Private OUTPUT__ As Variant

Private Sub Class_Initialize()
    INIT_
End Sub
Public Function Compile(src As String, Optional progName As String = "@anon") As Long
    If GLOBALS_ Is Nothing Then
        INIT_
    End If
    With Compiler_
        .SetGlobals GLOBALS_
        Compile = .CompileProgram(src, progName)
    End With
End Function
Public Function GetGlobals()
    Set GetGlobals = GLOBALS_
End Function
Private Sub INIT_()
    Set GLOBALS_ = New Globals
    GLOBALS_.ASF_InitGlobals
    Set Compiler_ = New Compiler
    Set VM_ = New VM
    VM_.Verbose = False
End Sub
Public Property Get OUTPUT_() As Variant
    vAssignment OUTPUT_, OUTPUT__
End Property
Public Sub Run(idx As Long)
    If Not GLOBALS_ Is Nothing Then
        With VM_
            .SetGlobals GLOBALS_
            .RunProgramByIndex idx
            vAssignment OUTPUT__, .OUTPUT_
                If .Verbose Then
                    Debug.Print "=== Runtime Log ==="
                    Dim elm As Variant
                    For Each elm In GLOBALS_.gRuntimeLog
                        Debug.Print elm
                    Next elm
                End If
        End With
    End If
End Sub
Private Sub vAssignment(ByRef var As Variant, ByRef vValue As Variant)
    If IsObject(vValue) Then
        Set var = vValue
    Else
        var = vValue
    End If
End Sub
Public Property Get Verbose() As Boolean
    Verbose = VM_.Verbose
End Property
Public Property Let Verbose(aValue As Boolean)
    VM_.Verbose = aValue
End Property
