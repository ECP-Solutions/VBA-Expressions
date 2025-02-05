VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ASF_STATEMENT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Base 0
Private P_ENTRY_POINT_ As Long
Private P_EXIT_POINT_ As Long ' End of blocks: }
Private P_AUX_EXIT_POINT_ As Long 'Defines the JUMP if it differs to EXIT_POINT_
Private P_ON_ERR_EXIT_POINT_ As Long
Private P_STACK_() As ASF_OPERATION 'List of operations: OPERATION_. Consider to use variant type
Private P_TYPE_ As S_TYPE
Private P_ASF_IN_TERM_ As Boolean
Private P_ARGUMENTS_TYPES_() As ARGUMENT_TYPE
Private P_RETURNS_ As Variant 'Returned value if any
Public Property Get ENTRY_POINT_() As Long
    ENTRY_POINT_ = P_ENTRY_POINT_
End Property
Public Property Let ENTRY_POINT_(pValue As Long)
    P_ENTRY_POINT_ = pValue
End Property
Public Property Get EXIT_POINT_() As Long
    EXIT_POINT_ = P_EXIT_POINT_
End Property
Public Property Let EXIT_POINT_(pValue As Long)
    P_EXIT_POINT_ = pValue
End Property
Public Property Get AUX_EXIT_POINT_() As Long
    AUX_EXIT_POINT_ = P_AUX_EXIT_POINT_
End Property
Public Property Let AUX_EXIT_POINT_(pValue As Long)
    P_AUX_EXIT_POINT_ = pValue
End Property
Public Property Get ON_ERR_EXIT_POINT_() As Long
    ON_ERR_EXIT_POINT_ = P_ON_ERR_EXIT_POINT_
End Property
Public Property Let ON_ERR_EXIT_POINT_(pValue As Long)
    P_ON_ERR_EXIT_POINT_ = pValue
End Property
Public Property Get STACK_(index As Long) As ASF_OPERATION
    Set STACK_ = P_STACK_(index)
End Property
Public Property Set STACK_(index As Long, pValue As ASF_OPERATION)
    Set P_STACK_(index) = pValue
End Property
Public Property Get TYPE_() As S_TYPE
    TYPE_ = P_TYPE_
End Property
Public Property Let TYPE_(pValue As S_TYPE)
    P_TYPE_ = pValue
End Property
Public Property Get ASF_IN_TERM_() As Boolean
    ASF_IN_TERM_ = P_ASF_IN_TERM_
End Property
Public Property Let ASF_IN_TERM_(pValue As Boolean)
    P_ASF_IN_TERM_ = pValue
End Property
Public Property Get ARGUMENTS_TYPES_() As ARGUMENT_TYPE()
    ARGUMENTS_TYPES_ = P_ARGUMENTS_TYPES_
End Property
Public Property Let ARGUMENTS_TYPES_(pValue() As ARGUMENT_TYPE)
    P_ARGUMENTS_TYPES_ = pValue
End Property
Public Property Get RETURNS_() As Variant
    RETURNS_ = P_RETURNS_
End Property
Public Property Let RETURNS_(pValue As Variant)
    P_RETURNS_ = pValue
End Property
Public Sub InitializeStack(stackSize As Long)
    Dim i As Long
    
    ReDim P_STACK_(0 To stackSize - 1)
    For i = 0 To stackSize - 1
        Set P_STACK_(i) = New ASF_OPERATION
    Next i
End Sub
Public Sub Execute(cScope As VBAexpressionsScope)
    Dim i As Long
    
    Select Case Me.TYPE_
    Case S_TYPE.ASSIGNMENT_
        P_STACK_(0).SCOPE_.CopyScope cScope
        Set P_STACK_(0).expression_.EvalScope = P_STACK_(0).SCOPE_
        P_STACK_(0).RETURNS_ = P_STACK_(0).expression_.Eval()
        P_STACK_(1).SCOPE_.CopyScope P_STACK_(0).expression_.EvalScope
        For i = 0 To UBound(P_STACK_(1).arguments_)
            P_STACK_(1).SCOPE_.VarValue(P_STACK_(1).arguments_(i)) = CStr(P_STACK_(0).RETURNS_)
        Next i
    End Select
End Sub
