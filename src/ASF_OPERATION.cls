VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ASF_OPERATION"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Base 0
Private P_arguments_ As Variant
Private P_ARGUMENTS_TYPES_() As ARGUMENT_TYPE
Private P_expression_ As VBAexpressions
Private P_function_id_ As String 'eval underlying function
Private P_level_ As Long
Private P_RETURNS_ As Variant
Private P_SCOPE_ As VBAexpressionsScope 'Operation variables resides here.
Private P_statements_() As ASF_STATEMENT   'Group of operations
Private P_TYPE_ As O_TYPE
Public Property Get arguments_() As Variant
    arguments_ = P_arguments_
End Property
Public Property Let arguments_(pValue As Variant)
    P_arguments_ = pValue
End Property
Public Property Get ARGUMENTS_TYPES_() As ARGUMENT_TYPE()
    ARGUMENTS_TYPES_ = P_ARGUMENTS_TYPES_
End Property
Public Property Let ARGUMENTS_TYPES_(pValue() As ARGUMENT_TYPE)
    P_ARGUMENTS_TYPES_ = pValue
End Property
Public Property Get expression_() As VBAexpressions
    Set expression_ = P_expression_
End Property
Public Property Set expression_(pValue As VBAexpressions)
    Set P_expression_ = pValue
End Property
Public Property Get function_id_() As String
    function_id_ = P_function_id_
End Property
Public Property Let function_id_(pValue As String)
    P_function_id_ = pValue
End Property
Public Property Get level_() As Long
    level_ = P_level_
End Property
Public Property Let level_(pValue As Long)
    P_level_ = pValue
End Property
Public Property Get RETURNS_() As Variant
    RETURNS_ = P_RETURNS_
End Property
Public Property Let RETURNS_(pValue As Variant)
    P_RETURNS_ = pValue
End Property
Public Property Get SCOPE_() As VBAexpressionsScope
    Set SCOPE_ = P_SCOPE_
End Property
Public Property Set SCOPE_(pValue As VBAexpressionsScope)
    Set P_SCOPE_ = pValue
End Property
Public Property Get statements_() As ASF_STATEMENT()
    statements_ = P_statements_
End Property
Public Property Let statements_(pValue() As ASF_STATEMENT)
    P_statements_ = pValue
End Property
Public Property Get TYPE_() As O_TYPE
    TYPE_ = P_TYPE_
End Property
Public Property Let TYPE_(pValue As O_TYPE)
    P_TYPE_ = pValue
End Property
Public Sub InitializeStack(stackSize As Long)
    Dim i As Long
    
    ReDim P_statements_(0 To stackSize - 1)
    For i = 0 To stackSize - 1
        Set P_statements_(i) = New ASF_STATEMENT
    Next i
End Sub
