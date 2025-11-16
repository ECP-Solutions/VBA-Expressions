VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
    Dim m As New Map
    m.Add "type", nodeType
    Set MakeNode = m
End Function

Public Sub RunProgramByIndex(idx As Long)
    GLOBALS_.ASF_InitGlobals
    If idx < 1 Or idx > GLOBALS_.gPrograms.Count Then Exit Sub
    Dim p As Variant: p = GLOBALS_.gPrograms(idx)
    Dim progName As String: progName = p(0)
    Dim stmts As Collection: Set stmts = p(1)
    Dim rawScope As Collection: Set rawScope = p(2)
    Dim progScope As New ScopeStack
    progScope.LoadRaw rawScope
    progScope.Push
    If VERBOSE_ Then GLOBALS_.gRuntimeLog.Add "RUN Program: " & progName
    Dim i As Long
    For i = 1 To stmts.Count
        Dim ctrl As String
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
            For i = 1 To args.Count
                Dim v As Variant: vAssignment v, EvalExprNode(args(i), progScope)
                outParts.Add ValueToStringForPrint(v)
            Next i
            Dim sb As String: sb = ""
            For i = 1 To outParts.Count
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
                If InStr(1, lName, ".") Then
                    Dim tLeft As Map
                    Dim tmpPropMap() As String
                    tmpPropMap() = Split(lName, ".")
                    Set tLeft = progScope.GetValue(tmpPropMap(LBound(tmpPropMap)))
                    For i = LBound(tmpPropMap) + 1 To UBound(tmpPropMap) - 1
                        Set tLeft = tLeft.GetValue(tmpPropMap(i))
                    Next i
                    tLeft.SetValue tmpPropMap(UBound(tmpPropMap)), rval
                End If
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

    arr = progScope.GetValue(arrName)
    pos = CLng(idxV)
    If pos < 1 Then
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
            ReDim Preserve arr(lb To lb + pos - 1)
        End If
    End If
    arr(LBound(arr) + pos - 1) = val
    progScope.SetValue arrName, arr
End Sub

' ---------------------
' ExecIfNode: Evaluate condition, then execute proper block
' ---------------------
Private Function ExecIfNode(node As Map, progScope As ScopeStack) As String
    Dim condNode As Map: Set condNode = node.GetValue("cond")
    If IsTruthy(EvalExprNode(condNode, progScope)) Then
        Dim thenStmts As Collection: Set thenStmts = node.GetValue("then")
        Dim si As Long
        For si = 1 To thenStmts.Count
            Dim ctrl As String: ctrl = ExecuteStmtNode(thenStmts(si), progScope)
            If ctrl <> "" Then ExecIfNode = ctrl: Exit Function
        Next si
        ExecIfNode = ""
        Exit Function
    End If
    Dim elseifConds As Collection: Set elseifConds = node.GetValue("elseif_conds")
    Dim elseifBlocks As Collection: Set elseifBlocks = node.GetValue("elseif_blocks")
    Dim idx As Long
    For idx = 1 To elseifConds.Count
        If IsTruthy(EvalExprNode(elseifConds(idx), progScope)) Then
            Dim bl As Collection: Set bl = elseifBlocks(idx)
            Dim si2 As Long
            For si2 = 1 To bl.Count
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
        For ee = 1 To els.Count
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
    For s = 1 To body.Count
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
    Dim m As Map: Set m = MakeNode("ExprStmt")
    m.SetValue "expr", expr
    Set MakeNodeExprStmt = m
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
    For i = 1 To body.Count
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
    For i = 1 To tryStmts.Count
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
        For j = 1 To catchStmts.Count
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
    For i = 1 To cases.Count
        Dim pair As Collection: Set pair = cases(i)
        Dim caseExpr As Map: Set caseExpr = pair(1)
        Dim blockStmts As Collection: Set blockStmts = pair(2)
        match = (EvalExprNode(caseExpr, progScope) = switchVal)
        If match Then
            Dim s As Long
            For s = 1 To blockStmts.Count
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
            For d = 1 To defBlk.Count
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
            Dim cmap As New Map
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
        Case "VBEXPR"
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
            For pi = 1 To items.Count
                Set pair = items(pi)
                key = CStr(pair(1))
                Set valNode = pair(2)
                om.SetValue key, EvalExprNode(valNode, progScope)
            Next pi
            ' Return Map object (as Variant holding the object)
            vAssignment tmpResult, om
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

            ' Defensive fallback: optional function-objects registry (NOT IMPLEMENTED YET).
'            On Error Resume Next
'            If Not (gFuncObjects Is Nothing) Then
'                If gFuncObjects.Exists(vName) Then
'                    EvalExprNode = gFuncObjects.GetValue(vName)
'                    On Error GoTo 0
'                    GoTo exitfun
'                End If
'            End If
'            On Error GoTo 0

            ' Not found — log to runtime log and return Empty
            If VERBOSE_ Then GLOBALS_.gRuntimeLog.Add "VM: Variable '" & vname & "' is undefined or not callable (returned Empty)."
            vAssignment tmpResult, Empty
            GoTo exitfun
        Case "Member"
            ' Member(base, prop) — evaluate base, then property access
            Set baseExpr = node.GetValue("base")
            Dim propStr As String: propStr = CStr(node.GetValue("prop"))
            Dim baseEvaluated As Variant
            vAssignment baseEvaluated, EvalExprNode(baseExpr, progScope)
            ' special-case array length: if prop == "length" and base is array, return length
            If propStr = "length" Then
                If IsArray(baseEvaluated) Then
                    Dim ub2 As Long
                    ub2 = UBound(baseEvaluated)
                    If ub2 < LBound(baseEvaluated) Then
                        vAssignment tmpResult, 0
                    Else
                        vAssignment tmpResult, ub2 - LBound(baseEvaluated) + 1
                    End If
                    GoTo exitfun
                End If
            End If
            ' if base is Map, return stored property
            If TypeName(baseEvaluated) = "Map" Then
                Dim bm As Map: Set bm = baseEvaluated
                vAssignment tmpResult, bm.GetValue(propStr)
                GoTo exitfun
            End If
            ' otherwise no property -> Empty
            vAssignment tmpResult, Empty
            GoTo exitfun
        Case "Array"
            Set items = node.GetValue("items")
            If items.Count = 0 Then
                vAssignment tmpResult, Array()
                GoTo exitfun
            End If
            Dim a() As Variant
            ReDim a(1 To items.Count)
            Dim ii As Long
            For ii = 1 To items.Count
                vAssignment a(ii), EvalExprNode(items(ii), progScope)
            Next ii
            vAssignment tmpResult, a
            GoTo exitfun
        Case "Index"
            Dim base As Map: Set base = node.GetValue("base")
            Dim idxNode As Map: Set idxNode = node.GetValue("index")
            If base.GetValue("type") = "Variable" Then
                Dim arr As Variant: arr = progScope.GetValue(base.GetValue("name"))
                If Not IsArray(arr) Then vAssignment tmpResult, Empty: GoTo exitfun
                Dim pos As Long: pos = CLng(EvalExprNode(idxNode, progScope))
                vAssignment tmpResult, arr(LBound(arr) + pos - 1)
                GoTo exitfun
            Else
                vAssignment tmpResult, Empty: GoTo exitfun
            End If
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
            For k = 1 To argsColl.Count
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
                        thisVal = EvalExprNode(baseExpr, progScope)
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
                rop = EvalExprNode(node.GetValue("right"), progScope)
                vAssignment tmpResult, (IsTruthy(lop) And IsTruthy(rop))
                GoTo exitfun
            ElseIf op2 = "||" Then
                If IsTruthy(lop) Then vAssignment tmpResult, True: GoTo exitfun
                rop = EvalExprNode(node.GetValue("right"), progScope)
                vAssignment tmpResult, (IsTruthy(lop) Or IsTruthy(rop))
                GoTo exitfun
            Else
                rop = EvalExprNode(node.GetValue("right"), progScope)
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
            If i - LBound(pa) + 1 <= args.Count Then
                callScope.SetValue CStr(pa(i)), args(i - LBound(pa) + 1)
            Else
                callScope.SetValue CStr(pa(i)), Empty
            End If
        Next i
    End If

    Dim stmts As Collection: Set stmts = p(1)
    For i = 1 To stmts.Count
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
                For i = 1 To visited.Count
                    If visited(i) Is v Then
                        ValueToStringWithCtx = "[Circular]": Exit Function
                    End If
                Next i
                visited.Add v
                ValueToStringWithCtx = MapToString(v, visited, depth + 1)
                visited.Remove visited.Count
                Exit Function

            Case "Collection"
                ' cycle detection
                Dim j As Long
                For j = 1 To visited.Count
                    If visited(j) Is v Then
                        ValueToStringWithCtx = "[Circular]": Exit Function
                    End If
                Next j
                visited.Add v
                ValueToStringWithCtx = CollectionToString(v, visited, depth + 1)
                visited.Remove visited.Count
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
Private Function MapToString(m As Variant, visited As Collection, depth As Long) As String
    On Error GoTo ErrHandler
    Dim keys As Collection: Set keys = m.keys
    Dim kcnt As Long: kcnt = keys.Count

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
            Dim val As Variant: val = m.GetValue(key)
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
        sb = sb & innerIndent & CStr(kk) & ": " & ValueToStringWithCtx(m.GetValue(CStr(kk)), visited, depth + 1)
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
    Dim n As Long: n = col.Count
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
    For i = 1 To visited.Count
        If visited(i) Is obj Then
            ObjectToString = "[Circular]": Exit Function
        End If
    Next i
    visited.Add obj
    ObjectToString = MapToString(obj, visited, depth)
    visited.Remove visited.Count
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
     Dim exprEval As VBAexpressions
     Set exprEval = New VBAexpressions

     ' Create expression in evaluator
     exprEval.Create expr

     ' Inject variables from progScope (shadowing: global frames first, then locals)
     Dim frame As Variant
     Dim m As Map
     Dim keyCol As Collection
     Dim key As Variant
     For Each frame In progScope.Raw
         Set m = frame
         Set keyCol = m.keys
         For Each key In keyCol
             exprEval.VarValue(CStr(key)) = m.GetValue(CStr(key))
         Next key
     Next frame

     ' Evaluate
     exprEval.Eval
     If exprEval.ErrorType = 0 Then
         EvalVBExpressionWithScope = exprEval.result
     Else
         ' On error, raise to caller; the try/catch at Exec layer can handle it
         err.Raise vbObjectError + 7001, "VM.EvalVBExpressionWithScope", "VBAexpressions eval error"
     End If
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
     Dim m As Map
     Dim keyCol As Collection
     Dim key As Variant
     For Each frame In progScope.Raw
         Set m = frame
         Set keyCol = m.keys
         For Each key In keyCol
             exprEval.VarValue(CStr(key)) = m.GetValue(CStr(key))
         Next key
     Next frame

     ' Inject arguments as temporary variables: __asf_vbarg_1, __asf_vbarg_2, ...
     Dim i As Long
     Dim tmpNames As New Collection
     For i = 1 To args.Count
         Dim tname As String: tname = "__asf_vbarg_" & CStr(i)
         tmpNames.Add tname
         exprEval.VarValue(tname) = args(i)
     Next i

     ' build call string referencing temp names
     Dim callStr As String: callStr = fname & "("
     For i = 1 To tmpNames.Count
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
             If i <= evaluatedArgs.Count Then
                 env.SetValue CStr(elm), evaluatedArgs(i)
             Else
                 env.SetValue CStr(elm), Empty
             End If
         Next elm
     End If

     ' set 'this' if provided
     If Not IsMissing(thisVal) Then
         env.SetValue "this", thisVal
     End If

     ' execute body
     Dim stmts As Collection: Set stmts = closureMap.GetValue("body")
     Dim ctrl As String
     Dim stmtIdx As Long
     For stmtIdx = 1 To stmts.Count
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
 Function ResolveLValue(lvalueNode As Map, progScope As ScopeStack) As Variant
     Dim result() As Variant
     If lvalueNode Is Nothing Then
         err.Raise vbObjectError + 9000, "VM.ResolveLValue", "Nil lvalue node"
     End If

     Dim t As String: t = ""
     If lvalueNode.Exists("type") Then t = CStr(lvalueNode.GetValue("type"))
     Dim idxVal2 As Variant
     
     Select Case t
     Case "Variable"
         Dim vname As String: vname = CStr(lvalueNode.GetValue("name"))
         ReDim result(0 To 3)
         result(0) = "scopeVar"
         result(1) = vname
         result(2) = vname
         ResolveLValue = result
         Exit Function

     Case "Member"
        ' member: compute/evaluate base container, ensure it is a Map if needed,
        ' and return ("mapObj", mapRef, propertyName)
         Dim propName As String
         If lvalueNode.Exists("name") Then
             propName = CStr(lvalueNode.GetValue("name"))
         ElseIf lvalueNode.Exists("property") Then
             propName = CStr(lvalueNode.GetValue("property"))
         Else
             err.Raise vbObjectError + 9001, "VM.ResolveLValue", "Member node missing property name"
         End If

         Dim baseNode As Map: Set baseNode = lvalueNode.GetValue("base")
        ' If base is a bare variable and empty, create Map and store it into the scope so further member sets work
        ' Special-case: if the base is an Index (array-element), we must obtain the actual
        ' element *stored in the array* (not a temporary value) so member writes mutate the stored element.
         If Not baseNode Is Nothing Then
             Dim baseTypeAhead As String
             If baseNode.Exists("type") Then baseTypeAhead = CStr(baseNode.GetValue("type")) Else baseTypeAhead = ""
             If baseTypeAhead = "Index" Then
                 ' Resolve the index target to the actual element Map (creating it if necessary)
                 Dim elemMap As Map: Set elemMap = GetElementMapFromIndexNode(baseNode, progScope)
                 If elemMap Is Nothing Then
                     err.Raise vbObjectError + 9020, "VM.ResolveLValue", "Index element not available for member assignment"
                 End If
                 ' Build a 4-element descriptor so HandleAssignment will find mapObj at resolved(1) and prop at resolved(3)
                 Dim outArrSpecial As Variant: outArrSpecial = Array("mapObj", elemMap, "", propName)
                 ResolveLValue = outArrSpecial
                 Exit Function
             End If
             Dim baseType As String
             If baseNode.Exists("type") Then baseType = CStr(baseNode.GetValue("type")) Else baseType = ""
             If baseType = "Variable" Then
                 Dim baseVarName As String: baseVarName = CStr(baseNode.GetValue("name"))
                 Dim baseVal As Variant: baseVal = progScope.GetValue(baseVarName)
                 If IsEmpty(baseVal) Then
                     Dim newMap As Map: Set newMap = MakeNode("Map")
                     progScope.SetValue baseVarName, newMap
                     baseVal = newMap
                 End If
                 If TypeName(baseVal) = "Map" Then
                     ReDim result(0 To 3)
                     result(0) = "mapObj"
                     vAssignment result(1), baseVal
                     vAssignment result(2), propName
                     ResolveLValue = result
                     Exit Function
                 Else
                     err.Raise vbObjectError + 9002, "VM.ResolveLValue", "Cannot assign member on non-object base"
                 End If
             Else
                 Dim baseVal2 As Variant: baseVal2 = EvalExprNode(baseNode, progScope)
                 If TypeName(baseVal2) = "Map" Then
                     ReDim result(0 To 3)
                     result(0) = "mapObj"
                     vAssignment result(1), baseVal2
                     result(2) = propName
                     ResolveLValue = result
                     Exit Function
                 Else
                     err.Raise vbObjectError + 9003, "VM.ResolveLValue", "Cannot assign member on non-object base"
                 End If
             End If
         Else
             err.Raise vbObjectError + 9004, "VM.ResolveLValue", "Member node missing base"
         End If

     Case "Index"
         Dim idxNode As Map: Set idxNode = lvalueNode.GetValue("index")
         Dim baseIdxNode As Map: Set baseIdxNode = lvalueNode.GetValue("base")
         If baseIdxNode Is Nothing Then
             err.Raise vbObjectError + 9005, "VM.ResolveLValue", "Index node missing base"
         End If
         Dim baseIdxType As String: baseIdxType = ""
         If baseIdxNode.Exists("type") Then baseIdxType = CStr(baseIdxNode.GetValue("type"))

         If baseIdxType = "Variable" Then
             Dim arrVarName As String: arrVarName = CStr(baseIdxNode.GetValue("name"))
             Dim arrVal As Variant: arrVal = progScope.GetValue(arrVarName)
             If IsEmpty(arrVal) Then
                 Dim emptyArr() As Variant
                 ReDim emptyArr(0 To 0)
                 progScope.SetValue arrVarName, emptyArr
                 arrVal = progScope.GetValue(arrVarName)
             End If
             Dim idxVal As Variant: idxVal = EvalExprNode(idxNode, progScope)
             If Not IsNumeric(idxVal) Then err.Raise vbObjectError + 9006, "VM.ResolveLValue", "Array index must be numeric"
             ReDim result(0 To 3)
             result(0) = "arrayInScope"
             result(1) = arrVarName
             result(2) = CLng(idxVal)
             ResolveLValue = result
             Exit Function
        Else
            ' base is an expression (e.g., a.prop[...]) — if base is Member(baseObj, propName)
            ' we want the parent map object and property name so we can assign into the array stored in the property.
            If baseIdxNode.Exists("type") And CStr(baseIdxNode.GetValue("type")) = "Member" Then
                Dim parentNode As Map: Set parentNode = baseIdxNode.GetValue("base")
                If parentNode Is Nothing Then err.Raise vbObjectError + 9015, "VM.ResolveLValue", "Member base missing"
                ' Evaluate parent object (should yield the Map that holds the property)
                Dim parentVal As Variant: parentVal = EvalExprNode(parentNode, progScope)
                If IsEmpty(parentVal) Then
                    ' If the parent is a variable and is empty, create a Map to hold properties (match other behaviors)
                    If parentNode.Exists("type") And CStr(parentNode.GetValue("type")) = "Variable" Then
                        Dim newParentMap As Map: Set newParentMap = MakeNode("Map")
                        progScope.SetValue CStr(parentNode.GetValue("name")), newParentMap
                        parentVal = newParentMap
                    End If
                End If
                If TypeName(parentVal) <> "Map" Then
                    err.Raise vbObjectError + 9016, "VM.ResolveLValue", "Parent of member index must be an object (Map)"
                End If
                Dim innerPropName As String: innerPropName = CStr(baseIdxNode.GetValue("name"))
                idxVal2 = EvalExprNode(idxNode, progScope)
                If Not IsNumeric(idxVal2) Then err.Raise vbObjectError + 9007, "VM.ResolveLValue", "Array index must be numeric"
                result = Array("arrayInMap", parentVal, CLng(idxVal2), innerPropName)
                ResolveLValue = result
                Exit Function
            Else
                ' fallback for generic expression bases: evaluate base to a value
                Dim baseVal3 As Variant: baseVal3 = EvalExprNode(baseIdxNode, progScope)
                idxVal2 = EvalExprNode(idxNode, progScope)
                If Not IsNumeric(idxVal2) Then err.Raise vbObjectError + 9007, "VM.ResolveLValue", "Array index must be numeric"
                If TypeName(baseVal3) = "Map" Then
                    ' indexing a Map directly is not supported for numeric indices
                    err.Raise vbObjectError + 9008, "VM.ResolveLValue", "Cannot assign indexed member on non-variable/map base"
                ElseIf IsArray(baseVal3) Then
                    err.Raise vbObjectError + 9009, "VM.ResolveLValue", "Cannot assign into temporary array expression"
                Else
                    err.Raise vbObjectError + 9010, "VM.ResolveLValue", "Unsupported index base for assignment"
                End If
            End If
         End If

     Case Else
         err.Raise vbObjectError + 9011, "VM.ResolveLValue", "Unsupported LValue node type: " & t
     End Select

 End Function

 Sub HandleAssignment(lhs As Map, rhsValue As Variant, progScope As ScopeStack)
'     Dim resolved As Variant
'     On Error GoTo HandlerErr
'     resolved = ResolveLValue(lhs, progScope)
'     Dim kind As String: kind = CStr(resolved(0))
'     Select Case kind
'     Case "scopeVar"
'         Dim vname As String: vname = CStr(resolved(1))
'         progScope.SetValue vname, rhsValue
'     Case "mapObj"
'         Dim mapObj As Map: Set mapObj = resolved(1)
'         Dim propName2 As String: propName2 = CStr(resolved(2))
'         mapObj.SetValue propName2, rhsValue
'     Case "arrayInScope"
'         Dim arrName As String: arrName = CStr(resolved(1))
'         Dim idx As Long: idx = CLng(resolved(2))
'         AssignToArrayByName arrName, idx, rhsValue, progScope
'     Case "arrayInMap"
'         If UBound(resolved) >= 4 Then
'             Dim mapObj2 As Map: Set mapObj2 = resolved(1)
'             Dim idx2 As Long: idx2 = CLng(resolved(2))
'             Dim propName3 As String: propName3 = CStr(resolved(3))
'             Dim arrVal As Variant: arrVal = mapObj2.GetValue(propName3)
'             If IsEmpty(arrVal) Then
'                 Dim newA() As Variant
'                 ReDim newA(1 To 0)
'                 arrVal = newA
'             End If
'             Call AssignToArrayValueInPlace(arrVal, idx2, rhsValue)
'             mapObj2.SetValue propName3, arrVal
'         Else
'             err.Raise vbObjectError + 9110, "VM.HandleAssignment", "Malformed ResolveLValue arrayInMap result"
'         End If
'     Case Else
'         err.Raise vbObjectError + 9111, "VM.HandleAssignment", "Unhandled resolved LValue kind: " & kind
'     End Select
'     Exit Sub
    Dim resolved As Variant
    Dim propName2 As String
    On Error GoTo HandlerErr
    resolved = ResolveLValue(lhs, progScope)
    ' Support both returned shapes: 3-element (standard) or 4-element (arrayInMap with prop name)
    Dim kind As String: kind = CStr(resolved(0))
    Select Case kind
    Case "scopeVar"
        Dim vname As String: vname = CStr(resolved(1))
        ' If the scope variable name contains dots (e.g. "o.a") — try to resolve to a parent Map in scope.
        If InStr(vname, ".") > 0 Then
            Dim curName As String: curName = vname
            Dim done As Boolean: done = False
            Do While InStr(curName, ".") > 0 And Not done
                Dim lastDot As Long: lastDot = InStrRev(curName, ".")
                Dim parentName As String: parentName = left$(curName, lastDot - 1)
                Dim propName As String: propName = Mid$(curName, lastDot + 1)
                Dim parentVal As Variant: vAssignment parentVal, progScope.GetValue(parentName)
                If Not IsEmpty(parentVal) Then
                    If TypeName(parentVal) = "Map" Then
                        ' write into parentVal[propName]
                        parentVal.SetValue propName, rhsValue
                        done = True
                        Exit Do
                    ElseIf IsArray(parentVal) Then
                        ' parentVal is array stored in a top-level variable; assign into it if propName is a numeric index
                        If IsNumeric(propName) Then
                            Call AssignToArrayValueInPlace(parentVal, CLng(propName), rhsValue)
                            progScope.SetValue parentName, parentVal
                            done = True
                            Exit Do
                        End If
                    End If
                End If
                curName = parentName
            Loop
            If done Then Exit Sub
            ' If not resolvable, fallthrough and set top-level variable (will create variable named "o.a" which is undesirable)
        End If
        progScope.SetValue vname, rhsValue
    Case "mapObj"
        Dim mapObj As Map
        ' resolved(1) can be either a Map object or (unexpectedly) a string like "o.a" in some flows.
        If TypeName(resolved(1)) = "String" Then
            ' try to resolve dotted name to a Map + prop
            Dim dotted As String: dotted = CStr(resolved(1))
            If InStr(dotted, ".") > 0 Then
                Dim leftPart As String: leftPart = left$(dotted, InStrRev(dotted, ".") - 1)
                Dim rightPart As String: rightPart = Mid$(dotted, InStrRev(dotted, ".") + 1)
                Dim leftVal As Variant: vAssignment leftVal, progScope.GetValue(leftPart)
                If Not IsEmpty(leftVal) And TypeName(leftVal) = "Map" Then
                    Set mapObj = leftVal
                    propName2 = rightPart
                    mapObj.SetValue propName2, rhsValue
                    Exit Sub
                End If
            End If
            ' fallback: treat resolved(1) as string variable name
            Dim xname As String: xname = CStr(resolved(1))
            If Not IsEmpty(progScope.GetValue(xname)) Then
                If TypeName(progScope.GetValue(xname)) = "Map" Then
                    Set mapObj = progScope.GetValue(xname)
                    Dim propName2b As String: propName2b = CStr(resolved(2))
                    mapObj.SetValue propName2b, rhsValue
                    Exit Sub
                Else
                    err.Raise vbObjectError + 9130, "VM.HandleAssignment", "Cannot resolve mapObj from string: " & dotted
                End If
            End If
        Else
            Set mapObj = resolved(1)
            propName2 = CStr(resolved(2))
            mapObj.SetValue propName2, rhsValue
        End If
     Case "arrayInScope"
        Dim arrName As String: arrName = CStr(resolved(1))
        Dim idx As Long: idx = CLng(resolved(2))
        ' If arrName is dotted (like "o.a"), resolve to parent Map and property name, then operate on that property's array
        If InStr(arrName, ".") > 0 Then
            Dim curD As String: curD = arrName
            Dim resolvedOK As Boolean: resolvedOK = False
            Do While InStr(curD, ".") > 0 And Not resolvedOK
                Dim lastDot2 As Long: lastDot2 = InStrRev(curD, ".")
                Dim pName As String: pName = left$(curD, lastDot2 - 1)
                Dim propN As String: propN = Mid$(curD, lastDot2 + 1)
                Dim pVal As Variant: vAssignment pVal, progScope.GetValue(pName)
                If Not IsEmpty(pVal) And TypeName(pVal) = "Map" Then
                    Dim arrVal2 As Variant: vAssignment arrVal2, pVal.GetValue(propN)
                    If IsEmpty(arrVal2) Then
                        Dim tmpC() As Variant
                        ReDim tmpC(1 To 0)
                        arrVal2 = tmpC
                    End If
                    Call AssignToArrayValueInPlace(arrVal2, idx, rhsValue)
                    pVal.SetValue propN, arrVal2
                    resolvedOK = True
                    Exit Do
                End If
                curD = pName
            Loop
            If resolvedOK Then Exit Sub
            ' fallthrough to try as top-level arr var
        End If
        AssignToArrayByName arrName, idx, rhsValue, progScope
     Case "arrayInMap"
        ' in the special shape the resolver returned Array(kind,map,index,propName)
        If UBound(resolved) >= 4 Then
            Dim mapObj2 As Map: Set mapObj2 = resolved(1)
            Dim idx2 As Long: idx2 = CLng(resolved(2))
            Dim propName3 As String: propName3 = CStr(resolved(3))
            Dim arrVal As Variant: vAssignment arrVal, mapObj2.GetValue(propName3)
            ' if not array, create
            If IsEmpty(arrVal) Then
                Dim newA() As Variant
                ReDim newA(0 To 0)
                arrVal = newA
            End If
            ' Assign into the array stored in the map property (use 1-based semantics)
            Call AssignToArrayValueInPlace(arrVal, idx2, rhsValue)
            ' store back the possibly resized array
            mapObj2.SetValue propName3, arrVal
        Else
            err.Raise vbObjectError + 9110, "VM.HandleAssignment", "Malformed ResolveLValue arrayInMap result"
        End If
    End Select
    Exit Sub
HandlerErr:
     err.Raise err.Number, err.Source, err.Description
 End Sub

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
     arrVal(pos) = newValue
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
