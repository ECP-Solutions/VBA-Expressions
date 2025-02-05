VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ASF_PROGRAM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Base 0
Private P_ID_ As String
Private P_STACK_() As ASF_OPERATION  'List of operations/statements
Private P_SCOPE_ As VBAexpressionsScope
Private P_RETURNS_ As Variant 'Returned value if any
Public Enum O_TYPE
    variable_set = 0
    expression_eval = 1
    function_eval = 2
    STATEMENT = 3
End Enum
Public Enum ARGUMENT_TYPE
    T_VAR = 0
    T_LIST = 1
    T_ARRAY = 2
End Enum
Public Enum S_TYPE
    ASSIGNMENT_ = 0
    IF_ = 1
    ELSE_ = 2
    ELSE_IF_ = 3
    FOR_LOOP_ = 4
    WHILE_LOOP_ = 5
    TRY_CATCH_ = 6
End Enum
Public Property Get ID_() As String
    ID_ = P_ID_
End Property
Public Property Let ID_(pValue As String)
    P_ID_ = pValue
End Property
Public Property Get STACK_() As ASF_OPERATION()
    STACK_ = P_STACK_
End Property
Public Property Let STACK_(pValue() As ASF_OPERATION)
    P_STACK_ = pValue
End Property
Public Property Get SCOPE_() As VBAexpressionsScope
    SCOPE_ = P_SCOPE_
End Property
Public Property Set SCOPE_(pValue As VBAexpressionsScope)
    Set P_SCOPE_ = pValue
End Property
Public Property Get RETURNS_() As Variant
    RETURNS_ = P_RETURNS_
End Property
Public Property Let RETURNS_(pValue As Variant)
    P_RETURNS_ = pValue
End Property
