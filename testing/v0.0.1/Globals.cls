VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Globals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Module: Globals
Option Explicit

Public gPrograms As Collection          ' each item: Variant array (name, stmts Collection, scope raw Collection)
Public gFuncTable As Map               ' name -> program index
Public gFuncParams As Map              ' name -> parameter name array
Public ExpressionCache As Map          ' expr string -> compiled RPN or cached info (optional)
Public ProgramCache As Map             ' source string -> program index
Public gRuntimeLog As Collection
Public DEBUG_TRACE As Boolean
Public STRICT_STRINGS As Boolean
Public gCallStack As Collection
Private MAX_CALL_DEPTH_ As Long

Public Function ASF_InitGlobals() As Globals
    If gPrograms Is Nothing Then Set gPrograms = New Collection
    If gFuncTable Is Nothing Then Set gFuncTable = New Map
    If gFuncParams Is Nothing Then Set gFuncParams = New Map
    If ExpressionCache Is Nothing Then Set ExpressionCache = New Map
    If ProgramCache Is Nothing Then Set ProgramCache = New Map
    If gRuntimeLog Is Nothing Then Set gRuntimeLog = New Collection
    If gCallStack Is Nothing Then Set gCallStack = New Collection
    DEBUG_TRACE = True
    STRICT_STRINGS = True
    MAX_CALL_DEPTH_ = 512
    Set ASF_InitGlobals = Me
End Function
Public Property Get MAX_CALL_DEPTH() As Long
    MAX_CALL_DEPTH = MAX_CALL_DEPTH_
End Property
Public Property Let MAX_CALL_DEPTH(aValue As Long)
    MAX_CALL_DEPTH_ = aValue
End Property
