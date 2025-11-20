VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ---------------------
' VM class
' ---------------------
' Class Module: VM (AST executor)
Option Explicit
Private Const VTP_MAX_DEPTH As Long = 8    ' safety depth limit
Private Const VTP_MAX_ITEMS_INLINE As Long = 8 ' prefer inline for small containers
Private VERBOSE_ As Boolean
Private GLOBALS_ As Globals
Private OUTPUT__ As Variant

' Executes AST nodes produced by Compiler (AST).
' Uses Map node types and ScopeStack. Logs to gRuntimeLog.

Public Sub SetGlobals(aGlobals As Globals)
    Set GLOBALS_ = aGlobals
End Sub
Public Property Get OUTPUT_() As Variant
    vAssignment OUTPUT_, OUTPUT__
End Property
Public Property Get Verbose() As Boolean
    Verbose = VERBOSE_
End Property
Public Property Let Verbose(aValue As Boolean)
    VERBOSE_ = aValue
End Property

' Node helpers (Map-based)
Private Function MakeNode(nodeType As String) As Map
    Dim M As New Map
    M.Add "type", nodeType
    Set MakeNode = M
End Function

Public Sub RunProgramByIndex(idx As Long)
    GLOBALS_.ASF_InitGlobals
    If idx < 1 Or idx > GLOBALS_.gPrograms.count Then Exit Sub
    Dim p As Variant: p = GLOBALS_.gPrograms(idx)
    Dim progName As String: progName = p(0)
    Dim stmts As Collection: Set stmts = p(1)
    Dim rawScope As Collection: Set rawScope = p(2)
    Dim optionBase As Long: optionBase = IIf(UBound(p) >= 3, p(3), 1)
    Dim progScope As New ScopeStack
    progScope.LoadRaw rawScope
    progScope.SetValue "__option_base", optionBase
    progScope.Push
    If VERBOSE_ Then GLOBALS_.gRuntimeLog.Add "RUN Program: " & progName
    Dim i As Long
    For i = 1 To stmts.count
        Dim ctrl As String
'        ctrl = ExecuteStmtNode(stmts(i), progScope)
        ctrl = ExecuteStmtNode(stmts(i), progScope)
        If ctrl = "RETURN" Then
            vAssignment OUTPUT__, progScope.GetValue("__return")
            Exit For
        End If
        If ctrl = "ERR" Then Exit For
    Next i
    progScope.Pop
End Sub

' Execute a statement node (Map). Return control signals: "", "BREAK", "CONTINUE", "RETURN", "ERR"
Private Function ExecuteStmtNode(node As Map, progScope As ScopeStack) As String
    On Error GoTo ErrHandler
    Dim tp As String: tp = node.GetValue("type")
    Dim rval As Variant
    Dim i As Long
    Select Case tp
        Case "Print"
            Dim args As Collection: Set args = node.GetValue("args")
            Dim outParts As New Collection
            For i = 1 To args.count
                Dim v As Variant: vAssignment v, EvalExprNode(args(i), progScope)
                outParts.Add ValueToStringForPrint(v)
            Next i
            Dim sb As String: sb = ""
            For i = 1 To outParts.count
                If i > 1 Then sb = sb & ", "
                sb = sb & outParts(i)
            Next i
            If VERBOSE_ Then GLOBALS_.gRuntimeLog.Add "PRINT:" & sb
            Debug.Print sb
            ExecuteStmtNode = ""
            Exit Function

        Case "Assign"
            Dim left As Map: Set left = node.GetValue("left")
            Dim right As Map: Set right = node.GetValue("right")
            vAssignment rval, EvalExprNode(right, progScope)
            HandleAssignment left, rval, progScope
            If left.GetValue("type") = "Variable" Then
                Dim lName As String: lName = left.GetValue("name")
            End If
            ExecuteStmtNode = ""
            Exit Function

        Case "ExprStmt"
            Dim res As Variant: res = EvalExprNode(node.GetValue("expr"), progScope)
            ExecuteStmtNode = ""
            Exit Function

        Case "If"
            ExecuteStmtNode = ExecIfNode(node, progScope)
            Exit Function

        Case "For"
            ExecuteStmtNode = ExecForNode(node, progScope)
            Exit Function

        Case "While"
            ExecuteStmtNode = ExecWhileNode(node, progScope)
            Exit Function

        Case "Break"
            ExecuteStmtNode = "BREAK": Exit Function
        Case "Continue"
            ExecuteStmtNode = "CONTINUE": Exit Function

        Case "Return"
            Dim rex As Map: Set rex = node.GetValue("expr")
            If Not rex Is Nothing Then vAssignment rval, EvalExprNode(rex, progScope) Else rval = Empty
            progScope.SetValue "__return", rval
            ExecuteStmtNode = "RETURN": Exit Function

        Case "TryCatch"
            ExecuteStmtNode = ExecTryCatchNode(node, progScope)
            Exit Function

        Case "Switch"
            ExecuteStmtNode = ExecSwitchNode(node, progScope)
            Exit Function

        Case Else
            ' unknown node type
            If VERBOSE_ Then GLOBALS_.gRuntimeLog.Add "Unknown statement node: " & tp
            ExecuteStmtNode = ""
            Exit Function
    End Select

ErrHandler:
    If VERBOSE_ Then GLOBALS_.gRuntimeLog.Add "VM statement error: " & err.Description
    err.Clear
    ExecuteStmtNode = "ERR"
End Function

' Reuse your AssignToArray logic
Private Sub AssignToArray(arrName As String, idxV As Variant, val As Variant, progScope As ScopeStack)
    Dim arr As Variant
    Dim pos As Long
    Dim ub As Long
    Dim lb As Long
    Dim optionBase As Long: optionBase = progScope.GetValue("__option_base")
    
    arr = progScope.GetValue(arrName)
    pos = CLng(idxV)
    If pos < optionBase Then
        err.Raise vbObjectError + 5001, "VM.AssignToArray", "Invalid array index (must be >=1)"
    End If
    If Not IsArray(arr) Then
        If IsEmpty(arr) Then
            arr = Array()
        Else
            err.Raise vbObjectError + 5002, "VM.AssignToArray", "Not an array"
        End If
    End If
    ub = -1
    If IsArray(arr) Then
        ub = UBound(arr)
    End If
    If ub < 0 Then
        ReDim arr(1 To pos)
    Else
        lb = LBound(arr)
        If pos > (ub - lb + 1) Then
            ReDim Preserve arr(lb To lb + (pos - optionBase))
        End If
    End If
    arr(LBound(arr) + (pos - optionBase)) = val
    progScope.SetValue arrName, arr
End Sub

Private Function EvalMemberNode(node As Map, progScope As ScopeStack) As Variant
    If Not node.Exists("prop") And Not node.Exists("property") Then err.Raise vbObjectError + 5010, "VM.EvalMemberNode", "Missing prop/property in Member node"
    Dim propKey As String: propKey = IIf(node.Exists("prop"), "prop", "property")  ' Twist: auto-detect
    Dim prop As String: prop = node.GetValue(propKey)
    Dim baseVal As Variant
    vAssignment baseVal, EvalExprNode(node.GetValue("base"), progScope)
    If TypeName(baseVal) <> "Map" Then err.Raise vbObjectError + 5007, "VM.EvalMemberNode", "Cannot access property on non-map"
    vAssignment EvalMemberNode, baseVal.GetValue(prop)
End Function

Private Function EvalIndexNode(node As Map, progScope As ScopeStack) As Variant
    If Not node.Exists("index") Then err.Raise vbObjectError + 5011, "VM.EvalIndexNode", "Missing index in Index node"
    Dim baseVal As Variant
    vAssignment baseVal, EvalExprNode(node.GetValue("base"), progScope)
    Dim idxV As Variant: vAssignment idxV, EvalExprNode(node.GetValue("index"), progScope)
    Dim pos As Long: pos = CLng(idxV)
    Dim optionBase As Long: optionBase = progScope.GetValue("__option_base")
    If pos < optionBase Then err.Raise vbObjectError + 5001, "VM.EvalIndexNode", "Invalid index (must >= base)"
    If IsArray(baseVal) Then
        Dim lb As Long: lb = LBound(baseVal)
        Dim ub As Long: ub = UBound(baseVal)
        Dim adjustedPos As Long: adjustedPos = lb + (pos - optionBase)
        If adjustedPos > ub Then err.Raise vbObjectError + 5005, "VM.EvalIndexNode", "Index out of bounds"
        vAssignment EvalIndexNode, baseVal(adjustedPos)
    ElseIf TypeName(baseVal) = "Map" Then
        vAssignment EvalIndexNode, baseVal.GetValue(pos)
    Else
        err.Raise vbObjectError + 5006, "VM.EvalIndexNode", "Cannot index non-array/non-map"
    End If
End Function

Private Function EvalArrayNode(node As Map, progScope As ScopeStack) As Variant
    Dim items As Collection: Set items = node.GetValue("items")
    Dim arr As Variant
    Dim optionBase As Long: optionBase = progScope.GetValue("__option_base")
    If items Is Nothing Or items.count = 0 Then
        ReDim arr(optionBase To optionBase)
    Else
        ReDim arr(optionBase To optionBase + items.count - 1)
        Dim ii As Long
        For ii = 1 To items.count
            vAssignment arr(optionBase + ii - 1), EvalExprNode(items(ii), progScope)
        Next ii
    End If
    vAssignment EvalArrayNode, arr
End Function

Private Function ParsePath(ByVal path As String) As Collection
    Dim col As New Collection
    Dim pos As Long: pos = 1
    Dim lenP As Long: lenP = Len(path)
    Dim currentPart As String
    While pos <= lenP
        Dim ch As String: ch = Mid(path, pos, 1)
        If ch = "." Then
            If currentPart <> "" Then col.Add currentPart
            currentPart = ""
            pos = pos + 1
        ElseIf ch = "[" Then
            If currentPart <> "" Then col.Add currentPart
            currentPart = ""
            Dim startB As Long: startB = pos
            pos = pos + 1
            While pos <= lenP And Mid(path, pos, 1) <> "]"
                pos = pos + 1
            Wend
            If pos > lenP Then err.Raise vbObjectError + 5003, "VM.ParsePath", "Unmatched bracket in path"
            currentPart = Mid(path, startB, pos - startB + 1)
            col.Add currentPart
            currentPart = ""
            pos = pos + 1
        Else
            currentPart = currentPart & ch
            pos = pos + 1
        End If
    Wend
    If currentPart <> "" Then col.Add currentPart
    Set ParsePath = col
End Function

Private Function ResolvePath(node As Map, progScope As ScopeStack, Optional forAssign As Boolean = False) As Map
    Dim res As New Map
    Dim tp As String: tp = node.GetValue("type")
    If tp = "Variable" Then
        If (InStr(node.GetValue("name"), ".") > 0 Or InStr(node.GetValue("name"), "[") > 0) Then
            ' Fallback for flat names (if Compiler didn't nest)
            Dim path As String: path = node.GetValue("name")
            Dim parts As Collection: Set parts = ParsePath(path) ' Add ParsePath from previous patch
            Dim current As Variant
            vAssignment current, progScope.GetValue(CStr(parts(1)))
            Dim ii As Long
            For ii = 2 To parts.count - IIf(forAssign, 1, 0)
                Dim part As String: part = parts(ii)
                If left(part, 1) = "[" Then
                    Dim indexStr As String: indexStr = Mid(part, 2, Len(part) - 2)
                    Dim indexV As Variant: indexV = EvalVBExpressionWithScope(indexStr, progScope) ' Twist: use VBEXPR for fallback computed
                    current = GetIndexedValue(current, indexV, progScope)
                Else
                    If TypeName(current) = "Map" Then
                        vAssignment current, current.GetValue(part)
                    Else
                        err.Raise vbObjectError + 5007, "VM.ResolvePath", "Cannot access property on non-map"
                    End If
                End If
            Next ii
            If forAssign Then
                res.SetValue "container", current
                res.SetValue "key", parts(parts.count)
                res.SetValue "kind", IIf(left(parts(parts.count), 1) = "[", "index", "prop")
            Else
                res.SetValue "value", current
            End If
        Else
            If forAssign Then
                res.SetValue "container", progScope
                res.SetValue "key", node.GetValue("name")
                res.SetValue "kind", "scopeVar"
            Else
                res.SetValue "value", progScope.GetValue(node.GetValue("name"))
            End If
        End If
    Else
        ' Nested nodes (preferred)
        Select Case tp
            Case "Member"
                Dim baseRes As Map: Set baseRes = ResolvePath(node.GetValue("base"), progScope, forAssign)
                If forAssign Then
                    res.SetValue "container", baseRes("container")
                    res.SetValue "key", node.GetValue("prop")
                    res.SetValue "kind", "prop"
                    Dim pKey As String
                    Dim pKind As String
                    ' carry forward parent property/key info if present on the base resolution
                    If baseRes.Exists("key") Then
                        pKey = baseRes("key")
                        res.SetValue "parentKey", pKey
                    End If
                    If baseRes.Exists("kind") Then
                        pKind = baseRes("kind")
                        res.SetValue "parentKind", pKind
                    End If
                    res.SetValue "parentObj", baseRes
                Else
                    res.SetValue "value", EvalMemberNode(node, progScope)
                End If
            Case "Index"
                Dim baseRes2 As Map: Set baseRes2 = ResolvePath(node.GetValue("base"), progScope, forAssign)
                If forAssign Then
                    ' We need to preserve the parent container _and_ the parent property name (if any)
                    ' so that assignments to an indexed property (e.g. o.a[2]) can write back correctly.
                    res.SetValue "container", baseRes2("container")
                    res.SetValue "key", EvalExprNode(node.GetValue("index"), progScope)
                    res.SetValue "kind", "index"
                    ' carry forward parent property/key info if present on the base resolution
                    If baseRes2.Exists("key") Then res.SetValue "parentKey", baseRes2("key")
                    If baseRes2.Exists("kind") Then res.SetValue "parentKind", baseRes2("kind")
                    res.SetValue "parentObj", baseRes2
                Else
                    res.SetValue "value", EvalIndexNode(node, progScope)
                End If
        End Select
    End If
    Set ResolvePath = res
End Function

Private Function GetIndexedValue(baseVal As Variant, idxV As Variant, progScope As ScopeStack) As Variant
    Dim pos As Long: pos = CLng(idxV)
    Dim optionBase As Long: optionBase = progScope.GetValue("__option_base")
    If pos < optionBase Then err.Raise vbObjectError + 5001, "VM.GetIndexedValue", "Invalid index"
    If IsArray(baseVal) Then
        Dim lb As Long: lb = LBound(baseVal)
        Dim adjustedPos As Long: adjustedPos = lb + (pos - optionBase)
        If adjustedPos > UBound(baseVal) Then err.Raise vbObjectError + 5005, "VM.GetIndexedValue", "Index out of bounds"
        vAssignment GetIndexedValue, baseVal(adjustedPos)
    ElseIf TypeName(baseVal) = "Map" Then
        vAssignment GetIndexedValue, baseVal.GetValue(pos)
    Else
        err.Raise vbObjectError + 5006, "VM.GetIndexedValue", "Cannot index non-array/non-map"
    End If
End Function

Private Function ResolveLValue(node As Map, progScope As ScopeStack) As Map
    Set ResolveLValue = ResolvePath(node, progScope, True)
    If node.GetValue("type") = "Index" Then
        If node.GetValue("index").GetValue("type") <> "Literal" Then
            ResolveLValue.Add "computed", True ' Twist: flag for lazy
        End If
    End If
End Function

' ---------------------
' ExecIfNode: Evaluate condition, then execute proper block
' ---------------------
Private Function ExecIfNode(node As Map, progScope As ScopeStack) As String
    Dim condNode As Map: Set condNode = node.GetValue("cond")
    If IsTruthy(EvalExprNode(condNode, progScope)) Then
        Dim thenStmts As Collection: Set thenStmts = node.GetValue("then")
        Dim si As Long
        For si = 1 To thenStmts.count
            Dim ctrl As String: ctrl = ExecuteStmtNode(thenStmts(si), progScope)
            If ctrl <> "" Then ExecIfNode = ctrl: Exit Function
        Next si
        ExecIfNode = ""
        Exit Function
    End If
    Dim elseifConds As Collection: Set elseifConds = node.GetValue("elseif_conds")
    Dim elseifBlocks As Collection: Set elseifBlocks = node.GetValue("elseif_blocks")
    Dim idx As Long
    For idx = 1 To elseifConds.count
        If IsTruthy(EvalExprNode(elseifConds(idx), progScope)) Then
            Dim bl As Collection: Set bl = elseifBlocks(idx)
            Dim si2 As Long
            For si2 = 1 To bl.count
                Dim ctrl2 As String: ctrl2 = ExecuteStmtNode(bl(si2), progScope)
                If ctrl2 <> "" Then ExecIfNode = ctrl2: Exit Function
            Next si2
            ExecIfNode = ""
            Exit Function
        End If
    Next idx
    If node.GetValue("hasElse") Then
        Dim els As Collection: Set els = node.GetValue("else")
        Dim ee As Long
        For ee = 1 To els.count
            Dim ctrl3 As String: ctrl3 = ExecuteStmtNode(els(ee), progScope)
            If ctrl3 <> "" Then ExecIfNode = ctrl3: Exit Function
        Next ee
    End If
    ExecIfNode = ""
End Function

' ---------------------
' ExecForNode
' ---------------------
Private Function ExecForNode(node As Map, progScope As ScopeStack) As String
    Dim initNode As Map: Set initNode = node.GetValue("init")
    Dim condNode As Map: Set condNode = node.GetValue("cond")
    Dim stepNode As Map: Set stepNode = node.GetValue("step")
    Dim body As Collection: Set body = node.GetValue("body")
    Dim condOk As Boolean
    
    If Not initNode Is Nothing Then
        Dim initType As String
        initType = initNode.GetValue("type")
        Select Case initType
            Case "Assign", "Print", "If", "For", "While", "TryCatch", "Switch", "Return", "Break", "Continue", "ExprStmt"
                ' Already a statement node — execute directly so side-effects happen.
                ExecuteStmtNode initNode, progScope
            Case Else
                ' Expression node — wrap into ExprStmt for side-effect evaluation.
                ExecuteStmtNode MakeNodeExprStmt(initNode), progScope
        End Select
    End If

ForStart:
    condOk = True
    If Not condNode Is Nothing Then
        condOk = IsTruthy(EvalExprNode(condNode, progScope))
    End If
    If Not condOk Then GoTo ForEnd

    Dim s As Long
    For s = 1 To body.count
        Dim ctrl As String: ctrl = ExecuteStmtNode(body(s), progScope)
        If ctrl = "BREAK" Then GoTo ForEnd
        If ctrl = "CONTINUE" Then Exit For
        If ctrl = "RETURN" Or ctrl = "ERR" Then ExecForNode = ctrl: Exit Function
    Next s

    If Not stepNode Is Nothing Then
        Dim stepType As String
        stepType = stepNode.GetValue("type")
        Select Case stepType
            Case "Assign", "Print", "If", "For", "While", "TryCatch", "Switch", "Return", "Break", "Continue", "ExprStmt"
                ExecuteStmtNode stepNode, progScope
            Case Else
                ExecuteStmtNode MakeNodeExprStmt(stepNode), progScope
        End Select
    End If
    GoTo ForStart

ForEnd:
    ExecForNode = ""
End Function

' Helper to wrap an expression node into an ExprStmt node
Private Function MakeNodeExprStmt(expr As Map) As Map
    Dim M As Map: Set M = MakeNode("ExprStmt")
    M.SetValue "expr", expr
    Set MakeNodeExprStmt = M
End Function

' ---------------------
' ExecWhileNode
' ---------------------
Private Function ExecWhileNode(node As Map, progScope As ScopeStack) As String
    Dim condNode As Map: Set condNode = node.GetValue("cond")
    Dim body As Collection: Set body = node.GetValue("body")

WhileStart:
    If Not IsTruthy(EvalExprNode(condNode, progScope)) Then GoTo WhileEnd
    Dim i As Long
    For i = 1 To body.count
        Dim ctrl As String: ctrl = ExecuteStmtNode(body(i), progScope)
        If ctrl = "BREAK" Then GoTo WhileEnd
        If ctrl = "CONTINUE" Then Exit For
        If ctrl = "RETURN" Or ctrl = "ERR" Then ExecWhileNode = ctrl: Exit Function
    Next i
    GoTo WhileStart

WhileEnd:
    ExecWhileNode = ""
End Function

' ---------------------
' ExecTryCatchNode
' ---------------------
Private Function ExecTryCatchNode(node As Map, progScope As ScopeStack) As String
    On Error GoTo TryErr
    Dim tryStmts As Collection: Set tryStmts = node.GetValue("try")
    Dim i As Long
    For i = 1 To tryStmts.count
        Dim ctrl As String: ctrl = ExecuteStmtNode(tryStmts(i), progScope)
        If ctrl = "RETURN" Or ctrl = "ERR" Then
            ExecTryCatchNode = ctrl
            If ctrl = "ERR" Then GoTo TryErr
            Exit Function
        End If
    Next i
    ExecTryCatchNode = ""
    Exit Function

TryErr:
    err.Clear
    Dim catchStmts As Collection: Set catchStmts = node.GetValue("catch")
    If Not catchStmts Is Nothing Then
        Dim j As Long
        For j = 1 To catchStmts.count
            Dim cctrl As String: cctrl = ExecuteStmtNode(catchStmts(j), progScope)
            If cctrl = "RETURN" Or cctrl = "ERR" Then ExecTryCatchNode = cctrl: Exit Function
        Next j
    End If
    ExecTryCatchNode = ""
End Function

' ---------------------
' ExecSwitchNode
' ---------------------
Private Function ExecSwitchNode(node As Map, progScope As ScopeStack) As String
    Dim switchVal As Variant: switchVal = EvalExprNode(node.GetValue("expr"), progScope)
    Dim cases As Collection: Set cases = node.GetValue("cases")
    Dim i As Long
    Dim match As Boolean
    For i = 1 To cases.count
        Dim pair As Collection: Set pair = cases(i)
        Dim caseExpr As Map: Set caseExpr = pair(1)
        Dim blockStmts As Collection: Set blockStmts = pair(2)
        match = (EvalExprNode(caseExpr, progScope) = switchVal)
        If match Then
            Dim s As Long
            For s = 1 To blockStmts.count
                Dim ctrl As String: ctrl = ExecuteStmtNode(blockStmts(s), progScope)
                If ctrl = "BREAK" Then ExecSwitchNode = "": Exit Function
                If ctrl = "RETURN" Or ctrl = "ERR" Then ExecSwitchNode = ctrl: Exit Function
            Next s
        End If
    Next i
    If Not match Then
        Dim defBlk As Collection: Set defBlk = node.GetValue("default")
        If Not defBlk Is Nothing Then
            Dim d As Long
            For d = 1 To defBlk.count
                Dim ctrl2 As String: ctrl2 = ExecuteStmtNode(defBlk(d), progScope)
                If ctrl2 = "BREAK" Then ExecSwitchNode = "": Exit Function
                If ctrl2 = "RETURN" Or ctrl2 = "ERR" Then ExecSwitchNode = ctrl2: Exit Function
            Next d
        End If
    End If
    ExecSwitchNode = ""
End Function

Private Sub vAssignment(ByRef var As Variant, ByRef vValue As Variant)
    If IsObject(vValue) Then
        Set var = vValue
    Else
        var = vValue
    End If
End Sub
' ---------------------
' Expression evaluator: Evaluate Expr AST nodes to runtime values
' ---------------------
Private Function EvalExprNode(node As Map, progScope As ScopeStack) As Variant
    If node Is Nothing Then EvalExprNode = Empty: Exit Function
    Dim tp As String: tp = node.GetValue("type")
    Dim items As Collection
    Dim tmpResult As Variant
    Dim pi As Long
    Dim funcIdx As Long
    Dim baseExpr As Map
    
    Select Case tp
        Case "FuncLiteral", "AnonFunc"
            ' Create closure (capture env by reference for shared-write closures)
            Dim cparams As Collection: vAssignment cparams, node.GetValue("params")
            Dim cbody As Collection: vAssignment cbody, node.GetValue("body")
            Dim cmap As Map
            Set cmap = New Map
            Dim envCopy As ScopeStack
            cmap.Add "type", "Closure"
            cmap.SetValue "params", cparams
            cmap.SetValue "body", cbody
            ' capture current scope reference (shared-write)
            Set envCopy = New ScopeStack
            envCopy.LoadRawByRef progScope.RawByRef
            cmap.SetValue "env", envCopy
            vAssignment tmpResult, cmap
            GoTo exitfun
        Case "VBAexpr"
            ' Forced VBAexpression node (from @(...)). The node stores the raw VB expression string in "expr".
            Dim rawVB As String: rawVB = node.GetValue("expr")
            vAssignment tmpResult, EvalVBExpressionWithScope(rawVB, progScope)
            GoTo exitfun
        Case "Object"
            ' Build and return a Map containing the evaluated properties
            Set items = node.GetValue("items")
            Dim om As New Map
            Dim valNode As Map
            Dim pair As Collection
            Dim key As String
            For pi = 1 To items.count
                Set pair = items(pi)
                key = CStr(pair(1))
                Set valNode = pair(2)
                om.SetValue key, EvalExprNode(valNode, progScope)
            Next pi
            ' Return Map object (as Variant holding the object)
            vAssignment tmpResult, om
            GoTo exitfun
        Case "Member"
            vAssignment tmpResult, EvalMemberNode(node, progScope)
            GoTo exitfun
        Case "Index"
            vAssignment tmpResult, EvalIndexNode(node, progScope)
            GoTo exitfun
        Case "Array"
            vAssignment tmpResult, EvalArrayNode(node, progScope)
            GoTo exitfun
        Case "Literal"
            vAssignment tmpResult, node.GetValue("value")
            GoTo exitfun
        Case "Variable"
            ' Robust variable resolution:
            '  - return actual value from scope (covers closures stored as Map AST nodes)
            '  - if scope has no value, fallback to named functions table (gFuncTable)
            '  - optional fallback to gFuncObjects registry
            Dim vname As String: vname = CStr(node.GetValue("name"))
            Dim vVal As Variant
            vAssignment vVal, progScope.GetValue(vname)

            ' If scope contains a non-empty value, return it — this covers closures
            ' stored as Map/AST nodes (FuncLiteral/Closure/etc.).
            If Not IsEmpty(vVal) Then
                ' If it's a Map and looks like a function node, return as callable
                On Error Resume Next
                If TypeName(vVal) = "Map" Then
                    Dim maybeType As Variant
                    maybeType = Empty
                    On Error Resume Next
                    maybeType = vVal.GetValue("type")
                    On Error GoTo 0
                    If Not IsEmpty(maybeType) Then
                        Dim tt As String: tt = CStr(maybeType)
                        ' Common names for function-like AST nodes / closures:
                        If tt = "FuncLiteral" Or tt = "Function" Or tt = "Closure" Or tt = "AnonFunction" Then
                            vAssignment tmpResult, vVal
                            GoTo exitfun
                        End If
                    End If
                End If
                ' Not necessarily a function-like Map — still return stored value.
                vAssignment tmpResult, vVal
                GoTo exitfun
            End If

            ' Not in scope (or Empty). Fallback: named compiled function table.
            On Error Resume Next
            If GLOBALS_.gFuncTable.Exists(vname) Then
                Dim fIdxVal As Variant: fIdxVal = GLOBALS_.gFuncTable.GetValue(vname)
                On Error GoTo 0
                If Not IsEmpty(fIdxVal) Then
                    funcIdx = CLng(fIdxVal)

                    ' Build a Closure map with the shape CallClosure expects:
                    ' "type" = "Closure", "params" = Collection, "body" = Collection, "env" = ScopeStack
                    Dim closureMap As New Map
                    closureMap.Add "type", "Closure"

                    ' params: convert stored gFuncParams (likely an array) into a Collection
                    Dim paramsCol As Collection
                    Set paramsCol = New Collection
                    If GLOBALS_.gFuncParams.Exists(vname) Then
                        Dim paVar As Variant: paVar = GLOBALS_.gFuncParams.GetValue(vname)
                        If IsArray(paVar) Then
                            If ArrayIsInit(paVar) Then
                                For pi = LBound(paVar) To UBound(paVar)
                                    paramsCol.Add CStr(paVar(pi))
                                Next pi
                            Else
                                paramsCol.Add vbNullString
                            End If
                        ElseIf TypeName(paVar) = "Collection" Then
                            ' already a collection — copy it
                            Dim it As Variant
                            For Each it In paVar
                                paramsCol.Add it
                            Next it
                        End If
                    End If
                    closureMap.SetValue "params", paramsCol

                    ' body and env come from gPrograms(funcIdx)
                    Dim pinfo As Variant
                    On Error Resume Next
                    pinfo = GLOBALS_.gPrograms(funcIdx)
                    If err.Number <> 0 Then
                        err.Clear
                        ' fallback: no program info — return Empty
                        EvalExprNode = Empty
                        Exit Function
                    End If
                    On Error GoTo 0

                    ' pinfo layout: Array(name, stmtsCollection, rawScopeCollection)
                    Dim bodyStmts As Collection
                    Set bodyStmts = pinfo(1)
                    closureMap.SetValue "body", bodyStmts

                    ' env: create ScopeStack and LoadRaw with stored raw scope (so closure has env object)
                    Dim envScope As ScopeStack
                    Set envScope = New ScopeStack
                    envScope.LoadRawByRef progScope.RawByRef
                    Set pinfo(2) = envScope.RawByRef
                    closureMap.SetValue "env", envScope

                    ' optionally store a name/funcIdx for debugging
                    closureMap.SetValue "name", vname
                    closureMap.SetValue "funcIdx", funcIdx

                    vAssignment tmpResult, closureMap
                    GoTo exitfun
                End If
            End If
            On Error GoTo 0

            ' Not found — log and return Empty
            If VERBOSE_ Then GLOBALS_.gRuntimeLog.Add "VM: Variable '" & vname & "' is undefined or not callable (returned Empty)."
            vAssignment tmpResult, Empty
            GoTo exitfun
        Case "Call"
            ' Call can be either:
            '  - Call with "name" (string) and args (legacy)
            '  - Call with "callee" (an expression node) and args
            '       (new: allows variable holding closures)

            Dim fname As String
            Dim hasCalleeExpr As Boolean
            If node.Exists("name") Then
                fname = CStr(node.GetValue("name"))
            End If
            If fname = "" Then
                ' try to derive name from callee expression when callee is a simple Variable node
                If node.Exists("callee") Then
                    hasCalleeExpr = True
                    Dim calleeExpr As Map: Set calleeExpr = node.GetValue("callee")
                    If Not calleeExpr Is Nothing Then
                        If calleeExpr.GetValue("type") = "Variable" Then
                            fname = CStr(calleeExpr.GetValue("name"))
                        End If
                    End If
                End If
            End If
            ' builtin length function
            If fname = ".__len__" Then
                Dim argn As Collection: Set argn = node.GetValue("args")
                Dim av As Variant: av = EvalExprNode(argn(1), progScope)
                If Not IsArray(av) Then
                    vAssignment tmpResult, 0
                Else
                    vAssignment tmpResult, UBound(av) - LBound(av) + 1
                End If
                GoTo exitfun
            End If

            Dim argsColl As Collection: Set argsColl = node.GetValue("args")
            Dim evaluated As Collection
            Set evaluated = New Collection
            Dim k As Long
            For k = 1 To argsColl.count
                evaluated.Add EvalExprNode(argsColl(k), progScope)
            Next k

            If Not hasCalleeExpr Then
                ' named call (existing behavior + fallback to VBAexpressions)
                If GLOBALS_.gFuncTable.Exists(fname) Then
                    funcIdx = CLng(GLOBALS_.gFuncTable.GetValue(fname))
                    vAssignment tmpResult, CallFuncByIndex_AST(funcIdx, evaluated)
                Else
                    ' named but not internal -> try VBAexpressions function call
                    vAssignment tmpResult, EvalVBFunctionCall(fname, evaluated, progScope)
                End If
            Else
                ' dynamic callee: evaluate callee expression -> it must yield a closure map or call via VB
                Dim calleeVal As Variant
                Dim hasThisVal As Boolean: hasThisVal = False
                Dim thisVal As Variant
                ' If callee is a member or index expression, compute base as `this` for method call binding.
                If Not calleeExpr Is Nothing Then
                    Dim ct As String: ct = IIf(calleeExpr.Exists("type"), CStr(calleeExpr.GetValue("type")), "")
                    If ct = "Member" Or ct = "Index" Then
                        hasThisVal = True
                        Set baseExpr = calleeExpr.GetValue("base")
                        vAssignment thisVal, EvalExprNode(baseExpr, progScope)
                    End If
                End If
                vAssignment calleeVal, EvalExprNode(calleeExpr, progScope)
                If TypeName(calleeVal) = "Map" Then
                    Dim cM As Map: Set cM = calleeVal
                    If cM.GetValue("type") = "Closure" Then
                        If hasThisVal Then
                            vAssignment tmpResult, CallClosure(cM, evaluated, thisVal)
                        Else
                            vAssignment tmpResult, CallClosure(cM, evaluated)
                        End If
                    Else
                        ' not a closure -> fail gracefully
                        vAssignment tmpResult, Empty
                    End If
                Else
                    ' if calleeVal is string -> call via VBAexpressions
                    If VarType(calleeVal) = vbString Then
                        vAssignment tmpResult, EvalVBFunctionCall(CStr(calleeVal), evaluated, progScope)
                    Else
                        vAssignment tmpResult, Empty
                    End If
                End If
            End If
            GoTo exitfun

        Case "Unary"
            Dim op As String: op = node.GetValue("op")
            Dim ev As Variant: ev = EvalExprNode(node.GetValue("expr"), progScope)
            If op = "!" Then
                vAssignment tmpResult, Not IsTruthy(ev)
            ElseIf op = "-" Then
                vAssignment tmpResult, -CDbl(ev)
            Else
                vAssignment tmpResult, ev
            End If
            GoTo exitfun
        Case "Ternary"
            Dim cnd As Variant: cnd = EvalExprNode(node.GetValue("cond"), progScope)
            If IsTruthy(cnd) Then
                vAssignment tmpResult, EvalExprNode(node.GetValue("trueExpr"), progScope)
            Else
                vAssignment tmpResult, EvalExprNode(node.GetValue("falseExpr"), progScope)
            End If
            GoTo exitfun
        Case "Binary"
            Dim lop As Variant: lop = EvalExprNode(node.GetValue("left"), progScope)
            Dim rop As Variant
            Dim op2 As String: op2 = node.GetValue("op")
            If op2 = "&&" Then
                If Not IsTruthy(lop) Then vAssignment tmpResult, False: GoTo exitfun
                vAssignment rop, EvalExprNode(node.GetValue("right"), progScope)
                vAssignment tmpResult, (IsTruthy(lop) And IsTruthy(rop))
                GoTo exitfun
            ElseIf op2 = "||" Then
                If IsTruthy(lop) Then vAssignment tmpResult, True: GoTo exitfun
                vAssignment rop, EvalExprNode(node.GetValue("right"), progScope)
                vAssignment tmpResult, (IsTruthy(lop) Or IsTruthy(rop))
                GoTo exitfun
            Else
                vAssignment rop, EvalExprNode(node.GetValue("right"), progScope)
                Select Case op2
                    Case "+"
                        vAssignment tmpResult, SafeAdd(lop, rop)
                    Case "-": vAssignment tmpResult, lop - rop
                    Case "*": vAssignment tmpResult, lop * rop
                    Case "/"
                        If rop = 0 Then err.Raise vbObjectError + 2001, "VM.EvalExprNode", "Division by zero" Else vAssignment tmpResult, lop / rop
                    Case "%": vAssignment tmpResult, lop Mod rop
                    Case "^": vAssignment tmpResult, lop ^ rop
                    Case "==", "=": vAssignment tmpResult, (lop = rop)
                    Case "!=": vAssignment tmpResult, (lop <> rop)
                    Case "<": vAssignment tmpResult, (lop < rop)
                    Case ">": vAssignment tmpResult, (lop > rop)
                    Case "<=": vAssignment tmpResult, (lop <= rop)
                    Case ">=": vAssignment tmpResult, (lop >= rop)
                    Case Else
                        vAssignment tmpResult, Empty
                End Select
                GoTo exitfun
            End If
    End Select

    vAssignment tmpResult, Empty
exitfun:
    vAssignment EvalExprNode, tmpResult
End Function

Private Function ReturnCollection(ByRef aColl As Variant) As Collection
    Set ReturnCollection = aColl
End Function
Private Function ArrayIsInit(aArray As Variant) As Boolean
    Dim ub As Long
    ArrayIsInit = True
    On Error GoTo err_handler
    ub = UBound(aArray)
    Exit Function
err_handler:
    err.Clear
    ArrayIsInit = False
End Function
' Call function program by index (AST). "args" is Collection of evaluated arg values.
Private Function CallFuncByIndex_AST(funcIdx As Long, args As Collection) As Variant
    Dim p As Variant: p = GLOBALS_.gPrograms(funcIdx)
    Dim rawScope As Collection: Set rawScope = p(2)
    Dim callScope As New ScopeStack
    Dim i As Long
    callScope.LoadRaw rawScope
    callScope.Push
    Dim fname As String: fname = p(0)
    ' set params from gFuncParams
    If GLOBALS_.gFuncParams.Exists(fname) Then
        Dim pa As Variant: pa = GLOBALS_.gFuncParams.GetValue(fname)
        For i = LBound(pa) To UBound(pa)
            If i - LBound(pa) + 1 <= args.count Then
                callScope.SetLocalValue CStr(pa(i)), args(i - LBound(pa) + 1)
            Else
                callScope.SetLocalValue CStr(pa(i)), Empty
            End If
        Next i
    End If

    Dim stmts As Collection: Set stmts = p(1)
    For i = 1 To stmts.count
        Dim ctrl As String: ctrl = ExecuteStmtNode(stmts(i), callScope)
        If ctrl = "RETURN" Then
            CallFuncByIndex_AST = callScope.GetValue("__return")
            callScope.Pop
            Exit Function
        ElseIf ctrl = "ERR" Then
            err.Raise vbObjectError + 3000, "VM.CallFunc", "Error during function execution"
        End If
    Next i
    CallFuncByIndex_AST = Empty
    callScope.Pop
End Function

' ---------------------
' Utilities adapted from your previous VM
' ---------------------
Private Function IsTruthy(v As Variant) As Boolean
    If IsObject(v) Then IsTruthy = Not v Is Nothing: Exit Function
    If IsNull(v) Then IsTruthy = False: Exit Function
    If IsEmpty(v) Then IsTruthy = False: Exit Function
    If VarType(v) = vbBoolean Then IsTruthy = CBool(v): Exit Function
    If IsNumeric(v) Then IsTruthy = (CDbl(v) <> 0): Exit Function
    If VarType(v) = vbString Then
        If IsBoolean(CStr(v)) Then IsTruthy = CBool(v) Else IsTruthy = (CStr(v) <> "")
        Exit Function
    End If
    IsTruthy = True
End Function

Private Function IsBoolean(ByRef expression As String) As Boolean
    IsBoolean = (LCase(expression) = "true")
    If Not IsBoolean Then IsBoolean = (LCase(expression) = "false")
End Function

Private Function SafeAdd(a As Variant, b As Variant) As Variant
    If IsNumeric(a) And IsNumeric(b) Then SafeAdd = a + b Else SafeAdd = CStr(a) & CStr(b)
End Function

'Private Function ValueToStringForPrint(v As Variant) As String
'    Dim ub As Long, s As String, i As Long
'    If Not IsArray(v) Then
'        If IsNull(v) Then ValueToStringForPrint = "NULL": Exit Function
'        If IsEmpty(v) Then ValueToStringForPrint = "": Exit Function
'        ' pretty-print Map objects as {k: v, ...}
'        If TypeName(v) = "Map" Then
'            Dim outMap As String: outMap = "{"
'            Dim keysCol As Collection
'            Set keysCol = v.keys
'            Dim kk As Long
'            For kk = 1 To keysCol.Count
'                If kk > 1 Then outMap = outMap & ", "
'                Dim kname As String: kname = CStr(keysCol(kk))
'                Dim kval As Variant: kval = v.GetValue(kname)
'                outMap = outMap & kname & ": " & ValueToStringForPrint(kval)
'            Next kk
'            outMap = outMap & "}"
'            ValueToStringForPrint = outMap
'            Exit Function
'        End If
'        ValueToStringForPrint = CStr(v): Exit Function
'    End If
'    ub = UBound(v)
'    If ub < 0 Then ValueToStringForPrint = "[]": Exit Function
'    s = "["
'    For i = LBound(v) To ub
'        If i > LBound(v) Then s = s & ", "
'        s = s & CStr(v(i))
'    Next i
'    ValueToStringForPrint = s & "]"
'End Function
Private Function ValueToStringForPrint(v As Variant) As String
    Dim visited As New Collection
    ValueToStringForPrint = ValueToStringWithCtx(v, visited, 0)
End Function
' core recursive printer with context
Private Function ValueToStringWithCtx(v As Variant, visited As Collection, depth As Long) As String
    On Error GoTo ErrHandler

    ' Depth guard
    If depth > VTP_MAX_DEPTH Then
        ValueToStringWithCtx = "..."
        Exit Function
    End If
    ' Null / Empty
    If IsNull(v) Then
        ValueToStringWithCtx = "NULL": Exit Function
    End If
    If IsEmpty(v) Then
        ValueToStringWithCtx = "": Exit Function
    End If

    ' Scalars
    If VarType(v) = vbString Then
        ValueToStringWithCtx = CStr(v): Exit Function
    End If
    If VarType(v) = vbBoolean Then
        If CBool(v) Then ValueToStringWithCtx = "True" Else ValueToStringWithCtx = "False"
        Exit Function
    End If
    If IsNumeric(v) Then
        ValueToStringWithCtx = CStr(v): Exit Function
    End If

    ' Arrays (native VBA arrays)
    If IsArray(v) Then
        ValueToStringWithCtx = ArrayToString(v, visited, depth)
        Exit Function
    End If

    ' Objects
    If IsObject(v) Then
        Dim tn As String: tn = TypeName(v)
        Select Case tn
            Case "Map"
                ' cycle detection by object identity
                Dim i As Long
                For i = 1 To visited.count
                    If visited(i) Is v Then
                        ValueToStringWithCtx = "[Circular]": Exit Function
                    End If
                Next i
                visited.Add v
                ValueToStringWithCtx = MapToString(v, visited, depth + 1)
                visited.Remove visited.count
                Exit Function

            Case "Collection"
                ' cycle detection
                Dim j As Long
                For j = 1 To visited.count
                    If visited(j) Is v Then
                        ValueToStringWithCtx = "[Circular]": Exit Function
                    End If
                Next j
                visited.Add v
                ValueToStringWithCtx = CollectionToString(v, visited, depth + 1)
                visited.Remove visited.count
                Exit Function

            Case Else
                ' generic objects: try to call a ToString-like property if present, else fallback
                ValueToStringWithCtx = ObjectToString(v, visited, depth + 1)
                Exit Function
        End Select
    End If

    ' Fallback
    ValueToStringWithCtx = CStr(v)
    Exit Function

ErrHandler:
    ' On any unexpected error, return a safe placeholder and continue
    On Error Resume Next
    ValueToStringWithCtx = "[error: " & err.Number & "]"
    err.Clear
End Function

' Convert Map -> string
Private Function MapToString(M As Variant, visited As Collection, depth As Long) As String
    On Error GoTo ErrHandler
    Dim keys As Collection: Set keys = M.keys
    Dim kcnt As Long: kcnt = keys.count

    If kcnt = 0 Then
        MapToString = "{}": Exit Function
    End If

    ' For small maps and shallow depth prefer inline representation
    If kcnt <= VTP_MAX_ITEMS_INLINE And depth <= 2 Then
        Dim parts() As String
        ReDim parts(1 To kcnt)
        Dim i As Long
        For i = 1 To kcnt
            Dim key As String: key = CStr(keys(i))
            Dim val As Variant: val = M.GetValue(key)
            parts(i) = CStr(key) & ": " & ValueToStringWithCtx(val, visited, depth)
        Next i
        MapToString = "{ " & Join(parts, ", ") & " }"
        Exit Function
    End If

    ' Multi-line pretty print
    Dim sb As String
    Dim indent As String: indent = String(depth * 2, " ")
    Dim innerIndent As String: innerIndent = String((depth + 1) * 2, " ")
    sb = "{"
    Dim first As Boolean: first = True
    Dim kk As Variant
    For Each kk In keys
        If Not first Then sb = sb & vbCrLf
        sb = sb & innerIndent & CStr(kk) & ": " & ValueToStringWithCtx(M.GetValue(CStr(kk)), visited, depth + 1)
        first = False
    Next kk
    sb = sb & vbCrLf & indent & "}"
    MapToString = sb
    Exit Function

ErrHandler:
    MapToString = "{<error>}"
    err.Clear
End Function

' Convert Collection -> string (treat as list)
Private Function CollectionToString(col As Variant, visited As Collection, depth As Long) As String
    On Error GoTo ErrHandler
    Dim n As Long: n = col.count
    If n = 0 Then CollectionToString = "[]": Exit Function
    If n <= VTP_MAX_ITEMS_INLINE And depth <= 2 Then
        Dim tmp() As String: ReDim tmp(1 To n)
        Dim ii As Long
        For ii = 1 To n
            tmp(ii) = ValueToStringWithCtx(col(ii), visited, depth)
        Next ii
        CollectionToString = "[ " & Join(tmp, ", ") & " ]"
        Exit Function
    End If

    Dim sb As String: sb = "["
    Dim i As Long
    Dim indent As String: indent = String((depth + 1) * 2, " ")
    For i = 1 To n
        If i > 1 Then sb = sb & vbCrLf
        sb = sb & indent & ValueToStringWithCtx(col(i), visited, depth + 1)
    Next i
    sb = sb & vbCrLf & String(depth * 2, " ") & "]"
    CollectionToString = sb
    Exit Function

ErrHandler:
    CollectionToString = "[<error>]"
    err.Clear
End Function

' Convert native VBA array -> string
Private Function ArrayToString(arr As Variant, visited As Collection, depth As Long) As String
    On Error GoTo ErrHandler
    Dim lb As Long, ub As Long
    lb = LBound(arr): ub = UBound(arr)
    Dim n As Long: n = ub - lb + 1
    If n <= 0 Then ArrayToString = "[]": Exit Function
    If n <= VTP_MAX_ITEMS_INLINE And depth <= 2 Then
        Dim tmp() As String: ReDim tmp(1 To n)
        Dim i As Long
        For i = lb To ub
            tmp(i - lb + 1) = ValueToStringWithCtx(arr(i), visited, depth)
        Next i
        ArrayToString = "[ " & Join(tmp, ", ") & " ]"
        Exit Function
    End If

    Dim sb As String: sb = "["
    Dim indent As String: indent = String((depth + 1) * 2, " ")
    Dim ii As Long
    For ii = lb To ub
        If ii > lb Then sb = sb & vbCrLf
        sb = sb & indent & ValueToStringWithCtx(arr(ii), visited, depth + 1)
    Next ii
    sb = sb & vbCrLf & String(depth * 2, " ") & "]"
    ArrayToString = sb
    Exit Function

ErrHandler:
    ArrayToString = "[<error>]"
    err.Clear
End Function

' Generic object to string fallback:
' - If the object is a Map-like (has Keys and GetValue), will attempt to treat it as Map.
' - Else TypeName + simple to-string
Private Function ObjectToString(obj As Variant, visited As Collection, depth As Long) As String
    On Error GoTo Fallback
    Dim tn As String: tn = TypeName(obj)

    ' Attempt Map-like duck typing: presence of Keys and GetValue
    ' (use On Error to bail out if methods missing)
    Dim dummy As Collection
    Dim tryKeys As Collection
    On Error GoTo Fallback2
    Set tryKeys = obj.keys
    ' if successful, treat as Map
    Dim i As Long
    For i = 1 To visited.count
        If visited(i) Is obj Then
            ObjectToString = "[Circular]": Exit Function
        End If
    Next i
    visited.Add obj
    ObjectToString = MapToString(obj, visited, depth)
    visited.Remove visited.count
    Exit Function

Fallback2:
    ' Not a Map-like object: try default string
    On Error GoTo Fallback
    ObjectToString = "<" & tn & ">"
    Exit Function

Fallback:
    ObjectToString = "<object>"
    err.Clear
End Function

' Utility to escape short strings for printing if you want (optional)
Private Function EscapeStringForPrint(s As String) As String
    ' currently returns s raw; adapt if you want quoted output
    EscapeStringForPrint = s
End Function
 ' VBAexpressions integration helpers
 ' ---------------------
 ' Evaluate a raw VBAexpressions expression string using a VBAexpressions instance,
 ' seeding it with the current ASF scope variables so VB expressions can reference ASF variables.
Private Function EvalVBExpressionWithScope(expr As String, progScope As ScopeStack) As Variant
     On Error GoTo ErrHandler
     With GLOBALS_.gExprEvaluator
         ' Create expression in evaluator
         .Create expr
    
         ' Inject variables from progScope (shadowing: global frames first, then locals)
         Dim frame As Variant
         Dim M As Map
         Dim keyCol As Collection
         Dim key As Variant
         For Each frame In progScope.Raw
             Set M = frame
             Set keyCol = M.keys
             For Each key In keyCol
                 .VarValue(CStr(key)) = M.GetValue(CStr(key))
             Next key
         Next frame
    
         ' Evaluate
         .Eval
         If .ErrorType = 0 Then
             EvalVBExpressionWithScope = .result
         Else
             ' On error, raise to caller; the try/catch at Exec layer can handle it
             err.Raise vbObjectError + 7001, "VM.EvalVBExpressionWithScope", "VBAexpressions eval error"
         End If
     End With
     Exit Function
ErrHandler:
     ' convert to runtime log and return Empty
     If VERBOSE_ Then GLOBALS_.gRuntimeLog.Add "VBAexpr error: " & err.Description
     err.Clear
     EvalVBExpressionWithScope = Empty
End Function

 ' Evaluate a function call in VBAexpressions. Args is a Collection of evaluated values.
 ' We create a temporary variable for each argument inside the VBAexpressions environment
 ' to avoid needing to serialize complex values into textual literals.
Private Function EvalVBFunctionCall(fname As String, args As Collection, progScope As ScopeStack) As Variant
     On Error GoTo ErrHandler
     Dim exprEval As VBAexpressions
     Set exprEval = New VBAexpressions

     ' Seed evaluator with ASF scope variables
     Dim frame As Variant
     Dim M As Map
     Dim keyCol As Collection
     Dim key As Variant
     For Each frame In progScope.Raw
         Set M = frame
         Set keyCol = M.keys
         For Each key In keyCol
             exprEval.VarValue(CStr(key)) = M.GetValue(CStr(key))
         Next key
     Next frame

     ' Inject arguments as temporary variables: __asf_vbarg_1, __asf_vbarg_2, ...
     Dim i As Long
     Dim tmpNames As New Collection
     For i = 1 To args.count
         Dim tname As String: tname = "__asf_vbarg_" & CStr(i)
         tmpNames.Add tname
         exprEval.VarValue(tname) = args(i)
     Next i

     ' build call string referencing temp names
     Dim callStr As String: callStr = fname & "("
     For i = 1 To tmpNames.count
         If i > 1 Then callStr = callStr & ","
         callStr = callStr & tmpNames(i)
     Next i
     callStr = callStr & ")"

     ' Evaluate
     exprEval.Create callStr
     exprEval.Eval
     If exprEval.ErrorType = 0 Then
         EvalVBFunctionCall = exprEval.result
     Else
         err.Raise vbObjectError + 7002, "VM.EvalVBFunctionCall", "VBAexpressions function call error"
     End If
     Exit Function
ErrHandler:
     If VERBOSE_ Then GLOBALS_.gRuntimeLog.Add "VBAexpr function-call error: " & err.Description
     err.Clear
     EvalVBFunctionCall = Empty
End Function

 ' ---------------------
 ' Closure / function-value support (runtime helpers)
 ' ---------------------

 ' Create and call closures (closureMap is a Map with keys:
 '  "type" = "Closure", "params" = Collection, "body" = Collection (stmts), "env" = ScopeStack)
 Private Function CallClosure(closureMap As Map, evaluatedArgs As Collection, Optional thisVal As Variant) As Variant
     On Error GoTo ErrHandler
     Dim env As ScopeStack
     Set env = closureMap.GetValue("env") ' shared reference (shared-write semantics)

     ' push a new frame for this call
     env.Push
     Dim i As Long
     Dim pa As Variant
     Dim elm As Variant
     Dim fc As Boolean
     
     vAssignment pa, closureMap.GetValue("params")
     If IsObject(pa) Then
        fc = (Not pa Is Nothing)
     Else
        fc = Not IsEmpty(pa)
     End If
     If fc Then
         For Each elm In pa
             i = i + 1
             If i <= evaluatedArgs.count Then
                 env.SetLocalValue CStr(elm), evaluatedArgs(i)
             Else
                 env.SetLocalValue CStr(elm), Empty
             End If
         Next elm
     End If

     ' set 'this' if provided
     If Not IsMissing(thisVal) Then
         env.SetLocalValue "this", thisVal
     End If

     ' execute body
     Dim stmts As Collection: Set stmts = closureMap.GetValue("body")
     Dim ctrl As String
     Dim stmtIdx As Long
     For stmtIdx = 1 To stmts.count
         ctrl = ExecuteStmtNode(stmts(stmtIdx), env)
         If ctrl = "RETURN" Then
             vAssignment CallClosure, env.GetValue("__return")
             env.Pop
             Exit Function
         ElseIf ctrl = "ERR" Then
             err.Raise vbObjectError + 8001, "VM.CallClosure", "Error during closure execution"
         End If
     Next stmtIdx

     ' normal return -> Empty
     env.Pop
     CallClosure = Empty
     Exit Function
ErrHandler:
     If VERBOSE_ Then GLOBALS_.gRuntimeLog.Add "CallClosure error: " & err.Description
     err.Clear
     CallClosure = Empty
 End Function
 ' ----------------- LValue resolution and helpers -----------------

Private Sub HandleAssignment(left As Map, ByVal rval As Variant, progScope As ScopeStack)
    Dim lval As Map: Set lval = ResolveLValue(left, progScope)
    Dim key As Variant: vAssignment key, lval("key")
    If lval.Exists("computed") Then vAssignment key, EvalExprNode(left.GetValue("index"), progScope) ' Lazy compute
    Dim container As Variant: vAssignment container, lval("container")
    Select Case lval("kind")
        Case "scopeVar"
            progScope.SetValue CStr(key), rval
        Case "prop"
            If TypeName(container) <> "Map" Then
                Dim tValue As Variant
                vAssignment tValue, ResolveIndexProp(lval)
                tValue.SetValue CStr(key), rval
            Else
                container.SetValue CStr(key), rval
            End If
        Case "index"
            Dim pos As Long: pos = CLng(key)
            Dim optionBase As Long: optionBase = progScope.GetValue("__option_base")
            If pos < optionBase Then err.Raise vbObjectError + 5001, "VM.HandleAssignment", "Invalid index"
        
            ' If ResolvePath provided parentKey this means the array is stored under a property/variable:
            Dim parentKey As Variant
            If lval.Exists("parentKey") Then vAssignment parentKey, lval("parentKey") Else parentKey = Empty
        
            If Not IsEmpty(parentKey) Then
                ' parent container holds the array under parentKey (could be progScope or Map)
                Dim arrVal As Variant
                If TypeName(container) = "Map" Then
                    vAssignment arrVal, container.GetValue(CStr(parentKey))
                ElseIf TypeName(container) = "ScopeStack" Then
                    vAssignment arrVal, container.GetValue(CStr(parentKey))
                Else
                    ' fallback: maybe container itself is the array
                    vAssignment arrVal, container
                End If
        
                ' Ensure arrVal is an array we can index into (AssignToArrayValueInPlace will create/resize).
                Call AssignToArrayValueInPlace(arrVal, pos, rval)
        
                ' Write the updated array back into the parent container
                If TypeName(container) = "Map" Then
                    container.SetValue CStr(parentKey), arrVal
                ElseIf TypeName(container) = "ScopeStack" Then
                    container.SetValue CStr(parentKey), arrVal
                Else
                    ' if container itself was an array variable (rare), try to overwrite it:
                    container = arrVal
                End If
            Else
                ' container is the array itself
                If IsArray(container) Then
                    Call AssignToArrayValueInPlace(container, pos, rval)
                Else
                    err.Raise vbObjectError + 5002, "VM.HandleAssignment", "Cannot index non-array/non-map for set"
                End If
            End If
        Case "Member"
            Dim baseVal As Variant: vAssignment baseVal, EvalExprNode(left.GetValue("base"), progScope)
            If TypeName(baseVal) <> "Map" Then err.Raise vbObjectError + 5007, "VM.HandleAssignment", "Cannot set property on non-map"
            Dim propKey As String: propKey = IIf(left.Exists("prop"), "prop", "property")
            baseVal.SetValue left.GetValue(propKey), rval
    End Select
End Sub
Private Function ResolveIndexProp(ByRef lval) As Variant
    Dim baseVariable As String
    Dim pObject As Variant
    Dim keys As New Collection
    Dim i As Long
    vAssignment pObject, lval("parentObj")
    Do While pObject.Exists("parentObj")
        keys.Add pObject("key")
        vAssignment pObject, pObject("parentObj")
        If Not pObject.Exists("parentKey") Then
            Exit Do
        End If
    Loop
    baseVariable = pObject("key")
    Dim container As Variant: vAssignment container, lval("container")
    vAssignment pObject, container.GetValue(baseVariable)
    For i = keys.count To 2 Step -1
        vAssignment pObject, pObject.GetValue(keys(i))
    Next
    vAssignment ResolveIndexProp, pObject(keys(1))
End Function
 Sub AssignToArrayByName(arrVarName As String, idx As Long, newValue As Variant, progScope As ScopeStack)
     Dim arrVal As Variant: arrVal = progScope.GetValue(arrVarName)
     If IsEmpty(arrVal) Then
         Dim tmpA() As Variant
         ReDim tmpA(1 To 0)
         arrVal = tmpA
     End If
     Call AssignToArrayValueInPlace(arrVal, idx, newValue)
     progScope.SetValue arrVarName, arrVal
 End Sub

 Sub AssignToArrayValueInPlace(ByRef arrVal As Variant, idx As Long, newValue As Variant)
     If Not IsArray(arrVal) Then
         Dim na() As Variant
         ReDim na(1 To idx)
         na(idx) = newValue
         arrVal = na
         Exit Sub
     End If
     Dim lb As Long: lb = LBound(arrVal)
     Dim ub As Long: ub = UBound(arrVal)
     Dim pos As Long: pos = lb + idx - 1
     If pos < lb Then
         err.Raise vbObjectError + 9120, "VM.AssignToArrayValueInPlace", "Index out of bounds (below LBound)"
     End If
     If pos > ub Then
         ReDim Preserve arrVal(lb To pos)
     End If
     vAssignment arrVal(pos), newValue
 End Sub
'-----------------------------------------------------------------------
' GetElementMapFromIndexNode
' Given an Index AST node, resolve the array container and return the actual
' element Map stored at the index (creating a Map element if empty).
'-----------------------------------------------------------------------
Function GetElementMapFromIndexNode(indexNode As Map, progScope As ScopeStack) As Map
    Dim resolved As Variant
    resolved = ResolveLValue(indexNode, progScope)
    If IsEmpty(resolved) Then
        Set GetElementMapFromIndexNode = Nothing
        Exit Function
    End If
    Dim kind As String: kind = CStr(resolved(0))
    Select Case kind
    Case "arrayInScope"
        Dim arrName As String: arrName = CStr(resolved(1))
        Dim idx As Long: idx = CLng(resolved(2))
        Dim arrVal As Variant: arrVal = progScope.GetValue(arrName)
        If IsEmpty(arrVal) Then
            Dim tmpA() As Variant
            ReDim tmpA(1 To 0)
            arrVal = tmpA
        End If
        Dim lb As Long: lb = LBound(arrVal)
        Dim pos As Long: pos = lb + idx - 1
        If pos > UBound(arrVal) Then
            ' extend array so the slot exists
            ReDim Preserve arrVal(lb To pos)
        End If
        Dim elem As Variant: elem = arrVal(pos)
        If IsEmpty(elem) Then
            Dim nm As Map: Set nm = MakeNode("Map")
            arrVal(pos) = nm
            progScope.SetValue arrName, arrVal
            Set GetElementMapFromIndexNode = nm
            Exit Function
        ElseIf TypeName(elem) = "Map" Then
            Set GetElementMapFromIndexNode = elem
            Exit Function
        Else
            ' not a Map stored in the array slot
            err.Raise vbObjectError + 9021, "VM.GetElementMapFromIndexNode", "Array element is not an object"
        End If
    Case "arrayInMap"
        ' resolved shape: Array("arrayInMap", mapObj, index, propName)
        Dim mapObj As Map: Set mapObj = resolved(1)
        Dim theIdx As Long: theIdx = CLng(resolved(2))
        Dim propName As String: propName = CStr(resolved(3))
        Dim arrVal2 As Variant: arrVal2 = mapObj.GetValue(propName)
        If IsEmpty(arrVal2) Then
            Dim tmpB() As Variant
            ReDim tmpB(1 To 0)
            arrVal2 = tmpB
        End If
        Dim lb2 As Long: lb2 = LBound(arrVal2)
        Dim pos2 As Long: pos2 = lb2 + theIdx - 1
        If pos2 > UBound(arrVal2) Then
            ReDim Preserve arrVal2(lb2 To pos2)
        End If
        Dim elem2 As Variant: elem2 = arrVal2(pos2)
        If IsEmpty(elem2) Then
            Dim nm2 As Map: Set nm2 = MakeNode("Map")
            arrVal2(pos2) = nm2
            mapObj.SetValue propName, arrVal2
            Set GetElementMapFromIndexNode = nm2
            Exit Function
        ElseIf TypeName(elem2) = "Map" Then
            Set GetElementMapFromIndexNode = elem2
            Exit Function
        Else
            err.Raise vbObjectError + 9022, "VM.GetElementMapFromIndexNode", "Array element is not an object (map)"
        End If
    Case Else
        err.Raise vbObjectError + 9023, "VM.GetElementMapFromIndexNode", "Unsupported resolved kind for index -> " & kind
    End Select
End Function
 ' ----------------- end LValue helpers -----------------
' ---------------------
' End VM class
' ---------------------

