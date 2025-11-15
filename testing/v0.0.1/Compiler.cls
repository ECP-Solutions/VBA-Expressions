VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Compiler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Module: Compiler
Option Explicit
Private GLOBALS_ As Globals

' Produces AST (Maps) for program and functions.
' Relies on Parser.Tokenize, Map, ScopeStack and globals:
'   ProgramCache, gPrograms (Collection), gFuncTable (Map), gFuncParams (Map)

Public Sub SetGlobals(ByRef aGlobals As Globals)
    Set GLOBALS_ = aGlobals
End Sub

' Node helpers (Map-based)
Private Function MakeNode(nodeType As String) As Map
    Dim m As New Map
    m.Add "type", nodeType
    Set MakeNode = m
End Function

' ---------------------
' Utility: deep-clone a Collection of Maps or primitive values
' ---------------------
Private Function CloneCollectionOfVariants(src As Variant) As Collection
    Dim out As New Collection
    Dim it As Variant
    For Each it In src
        If IsObject(it) Then
            If TypeName(it) = "Map" Then
                out.Add it.Clone
            ElseIf TypeName(it) = "Collection" Then
                out.Add CloneCollectionOfVariants(it)
            Else
                out.Add it
            End If
        Else
            out.Add it
        End If
    Next it
    Set CloneCollectionOfVariants = out
End Function

' ---------------------

Public Function CompileProgram(src As String, Optional progName As String = "@anon") As Long
    Dim p As Parser
    Dim toks As Collection
    Dim progScope As ScopeStack
    Dim stmtsAST As Collection
    Dim stmtTokens As Collection
    Dim i As Long
    Dim node As Map
    
    GLOBALS_.ASF_InitGlobals

    If GLOBALS_.ProgramCache.Exists(src) Then
        CompileProgram = GLOBALS_.ProgramCache.GetValue(src)
        Exit Function
    End If

    Set p = New Parser
    With p
        .SetGlobals GLOBALS_
        Set toks = .Tokenize(src)
    End With
    
    Set progScope = New ScopeStack
    progScope.Push

    Set stmtsAST = New Collection
    i = 1

Compiler_MainLoop:
    Do While i <= toks.Count
        ' skip comments at top level
        If toks(i)(0) = "COMMENT" Then
            i = i + 1
            GoTo Compiler_MainLoop
        End If

        ' Function definition at top-level
        If toks(i)(0) = "IDENT" And LCase$(toks(i)(1)) = "fun" Then
            i = i + 1
            ' skip comments
            Do While i <= toks.Count And toks(i)(0) = "COMMENT"
                i = i + 1
            Loop
            If i > toks.Count Then err.Raise vbObjectError + 1, , "Unexpected end after fun"
            Dim fname As String: fname = toks(i)(1)
            i = i + 1
            Do While i <= toks.Count And toks(i)(0) = "COMMENT"
                i = i + 1
            Loop

            Dim argList As New Collection
            If i <= toks.Count And toks(i)(0) = "PAREN" And toks(i)(1) = "(" Then
                i = i + 1
                Do While i <= toks.Count And Not (toks(i)(0) = "PAREN" And toks(i)(1) = ")")
                    If toks(i)(0) = "COMMENT" Then
                        i = i + 1
                    Else
                        If toks(i)(0) = "IDENT" Then argList.Add toks(i)(1)
                        i = i + 1
                    End If
                Loop
                If i <= toks.Count And toks(i)(0) = "PAREN" And toks(i)(1) = ")" Then i = i + 1
            End If

            Do While i <= toks.Count And toks(i)(0) = "COMMENT"
                i = i + 1
            Loop
            If i > toks.Count Or Not (toks(i)(0) = "PAREN" And toks(i)(1) = "{") Then
                err.Raise vbObjectError + 2, , "Expected function body"
            End If
            i = i + 1
            Dim depth As Long: depth = 0
            Dim bodyTokens As New Collection
            Do While i <= toks.Count
                If toks(i)(0) = "PAREN" Then
                    If toks(i)(1) = "{" Then
                        depth = depth + 1
                    ElseIf toks(i)(1) = "}" Then
                        If depth = 0 Then
                            Exit Do
                        Else
                            depth = depth - 1
                        End If
                    End If
                End If
                bodyTokens.Add toks(i)
                i = i + 1
            Loop

            Dim funcIdx As Long
            funcIdx = CompileFunctionAST(fname, argList, bodyTokens)
            GLOBALS_.gFuncTable.Add fname, funcIdx

            Dim pa() As String
            If argList.Count > 0 Then ReDim pa(0 To argList.Count - 1)
            Dim j As Long
            For j = 1 To argList.Count: pa(j - 1) = argList(j): Next j
            GLOBALS_.gFuncParams.Add fname, pa

            If i <= toks.Count Then i = i + 1 ' skip the trailing '}' if present
            GoTo Compiler_MainLoop
        Else
            ' collect a top-level statement tokens (nesting-aware) then parse to AST node
            Dim braceDepth As Long: braceDepth = 0
            Dim parenDepth As Long: parenDepth = 0
            Dim bracketDepth As Long: bracketDepth = 0
            
            Set stmtTokens = New Collection
            
            Do While i <= toks.Count
                If toks(i)(0) = "COMMENT" Then
                    i = i + 1
                    GoTo CollectNextTop
                End If

                If toks(i)(0) = "PAREN" Then
                    Select Case toks(i)(1)
                        Case "{"
                            braceDepth = braceDepth + 1
                        Case "}"
                            If braceDepth > 0 Then
                                braceDepth = braceDepth - 1
                            End If
                        Case "("
                            parenDepth = parenDepth + 1
                        Case ")"
                            If parenDepth > 0 Then
                                parenDepth = parenDepth - 1
                            End If
                        Case "["
                            bracketDepth = bracketDepth + 1
                        Case "]"
                            If bracketDepth > 0 Then
                                bracketDepth = bracketDepth - 1
                            End If
                    End Select
                End If

                ' Top-level statement separator: semicolon (commas are argument separators) if we are not nested
                If toks(i)(0) = "SEP" And toks(i)(1) = ";" And braceDepth = 0 And parenDepth = 0 And bracketDepth = 0 Then
                    i = i + 1
                    Exit Do
                End If

                stmtTokens.Add toks(i)
                i = i + 1
CollectNextTop:
            Loop

            If stmtTokens.Count > 0 Then
                Set node = ParseStatementTokensToAST(stmtTokens)
                If Not node Is Nothing Then
                    ' add a clone to avoid accidental reference reuse
                    stmtsAST.Add node.Clone
                End If
            End If
            GoTo Compiler_MainLoop
        End If
    Loop
    Dim pIndex As Long
    Dim normalizedStmts As Collection
    ' Normalize assignment-like expression-statements into proper Assign AST nodes
    Set normalizedStmts = NormalizeAssignsInStmts(stmtsAST)
    ' Normalize compound assignments (+=, -=, etc.)
    Set normalizedStmts = NormalizeCompoundAssigns(normalizedStmts)
    pIndex = GLOBALS_.gPrograms.Count + 1
    GLOBALS_.gPrograms.Add Array(progName, normalizedStmts, progScope.Raw)
    GLOBALS_.ProgramCache.Add src, pIndex
    CompileProgram = pIndex
End Function

' Compile a function body (AST) and register it as separate program
Private Function CompileFunctionAST(fname As String, argList As Collection, bodyTokens As Collection) As Long
    Dim funcScope As ScopeStack
    Dim i As Long
    Dim bodyStmtsAST As Collection
    Set funcScope = New ScopeStack
    funcScope.Push
    For i = 1 To argList.Count
        funcScope.SetValue argList(i), Empty
    Next i
    Dim idx As Long
    Set bodyStmtsAST = ParseTokensToAST(bodyTokens)
    ' Normalize assignments inside function body as well
    Set bodyStmtsAST = NormalizeAssignsInStmts(bodyStmtsAST)
    idx = GLOBALS_.gPrograms.Count + 1
    GLOBALS_.gPrograms.Add Array(fname, bodyStmtsAST, funcScope.Raw)
    CompileFunctionAST = idx
End Function

' Parse a list of tokens representing a block into AST statements (returns Collection of Map nodes)
Private Function ParseTokensToAST(toks As Collection) As Collection
    Dim out As Collection
    Dim stmtTokens As Collection
    Set out = New Collection
    Dim i As Long: i = 1
    Dim closedTopLevelBlock As Boolean
    
ParseMain:
    Do While i <= toks.Count
        If toks(i)(0) = "COMMENT" Then
            i = i + 1
            GoTo ParseMain
        End If
        Set stmtTokens = New Collection
        Dim braceDepth As Long: braceDepth = 0
        Dim parenDepth As Long: parenDepth = 0
        Dim bracketDepth As Long: bracketDepth = 0

ParseInner:
        Do While i <= toks.Count
            If toks(i)(0) = "COMMENT" Then
                i = i + 1
                GoTo ParseInner
            End If
            If toks(i)(0) = "PAREN" Then
                Select Case toks(i)(1)
                    Case "{"
                        braceDepth = braceDepth + 1
                    Case "}"
                        ' decrement only if we have an inner brace to close
                        If braceDepth > 0 Then
                            braceDepth = braceDepth - 1
                        End If
                        ' If we've just closed a top-level block (braceDepth now 0),
                        ' that usually signals the end of a statement (e.g. end of an if/for/while block).
                        ' Finalize the current statement tokens so adjacent statements are parsed separately.
                        If braceDepth = 0 Then
                            ' include this '}' token into the current stmtTokens and finish the statement
                            stmtTokens.Add toks(i)
                            i = i + 1
                            ' mark that we closed a top-level block for follow-up checks
                            closedTopLevelBlock = True
                            Exit Do
                        End If
                    Case "("
                        parenDepth = parenDepth + 1
                    Case ")"
                        If parenDepth > 0 Then
                            parenDepth = parenDepth - 1
                        End If
                    Case "["
                        bracketDepth = bracketDepth + 1
                    Case "]"
                        If bracketDepth > 0 Then
                            bracketDepth = bracketDepth - 1
                        End If
                End Select
            End If

            ' Only a semicolon (;) is a top-level statement separator.
            If toks(i)(0) = "SEP" And toks(i)(1) = ";" And braceDepth = 0 And parenDepth = 0 And bracketDepth = 0 Then
                i = i + 1
                Exit Do
            End If
    
            ' default: append current token and continue
            stmtTokens.Add toks(i)
            i = i + 1
        Loop

        If stmtTokens.Count > 0 Then
            Dim tmpNode As Map
            Set tmpNode = ParseStatementTokensToAST(stmtTokens)
            If Not tmpNode Is Nothing Then
                ' If we closed a top-level block and the next token is not a semicolon,
                ' warn the user (do not silently swallow).
                If closedTopLevelBlock Then
                    If i <= toks.Count Then
                        If Not (toks(i)(0) = "SEP" And toks(i)(1) = ";") Then
                            ' If the next token is not a comment and not end-of-input, log warning.
                            If Not (toks(i)(0) = "COMMENT") Then
                                On Error Resume Next
                                ' best-effort: add to runtime log if present, and print to Immediate
                                If Not GLOBALS_.gRuntimeLog Is Nothing Then GLOBALS_.gRuntimeLog.Add "COMPILER: missing ';' after block near token index " & CStr(i)
                                On Error GoTo 0
                            End If
                        Else
                            ' consume explicit semicolon
                            i = i + 1
                        End If
                    End If
                    closedTopLevelBlock = False
                End If
                out.Add tmpNode.Clone
            End If
        End If
        GoTo ParseMain
    Loop

    Set ParseTokensToAST = out
End Function

' Parse a single statement token collection into an AST node (Map)
Private Function ParseStatementTokensToAST(stmtTokens As Collection) As Map
    If stmtTokens.Count = 0 Then
        Set ParseStatementTokensToAST = Nothing
        Exit Function
    End If

    ' quick check for print
    If stmtTokens.Count >= 2 Then
        If stmtTokens(1)(0) = "IDENT" And LCase$(stmtTokens(1)(1)) = "print" Then
            Dim k As Long: k = 2
            Do While k <= stmtTokens.Count And Not (stmtTokens(k)(0) = "PAREN" And stmtTokens(k)(1) = "(")
                k = k + 1
            Loop
            If k <= stmtTokens.Count Then
                Dim opened As Long: opened = 1
                k = k + 1
                Dim innerToks As Collection
                Set innerToks = New Collection
                Do While k <= stmtTokens.Count And opened > 0
                    If stmtTokens(k)(0) = "PAREN" Then
                        If stmtTokens(k)(1) = "(" Then
                            opened = opened + 1
                        ElseIf stmtTokens(k)(1) = ")" Then
                            opened = opened - 1
                        End If
                    End If
                    If opened > 0 Then innerToks.Add stmtTokens(k)
                    k = k + 1
                Loop
                Dim args As Collection: Set args = ParseArgsTokensToExprNodes(innerToks)
                Dim node As Map: Set node = MakeNode("Print")
                node.SetValue "args", args
                Set ParseStatementTokensToAST = node
                Exit Function
            End If
        End If
    End If

    ' keywords
    If stmtTokens(1)(0) = "IDENT" Then
        Dim kw As String: kw = LCase$(stmtTokens(1)(1))
        Select Case kw
            Case "if"
                Set ParseStatementTokensToAST = ParseIfAST(stmtTokens)
                Exit Function
            Case "for"
                Set ParseStatementTokensToAST = ParseForAST(stmtTokens)
                Exit Function
            Case "while"
                Set ParseStatementTokensToAST = ParseWhileAST(stmtTokens)
                Exit Function
            Case "try"
                Set ParseStatementTokensToAST = ParseTryCatchAST(stmtTokens)
                Exit Function
            Case "switch"
                Set ParseStatementTokensToAST = ParseSwitchAST(stmtTokens)
                Exit Function
            Case "break"
                Dim b As Map: Set b = MakeNode("Break")
                Set ParseStatementTokensToAST = b
                Exit Function
            Case "continue"
                Dim c As Map: Set c = MakeNode("Continue")
                Set ParseStatementTokensToAST = c
                Exit Function
            Case "return"
                Dim rv As Map: Set rv = MakeNode("Return")
                If stmtTokens.Count >= 2 Then
                    Dim rhs As New Collection, m As Long
                    For m = 2 To stmtTokens.Count: rhs.Add stmtTokens(m): Next m
                    Dim rexpr As Map: Set rexpr = ParseExprTokensToNode(rhs)
                    rv.SetValue "expr", rexpr
                Else
                    rv.SetValue "expr", Nothing
                End If
                Set ParseStatementTokensToAST = rv
                Exit Function
        End Select
    End If

    ' assignment: find top-level "=" operator position
    Dim assignPos As Long: assignPos = 0
    Dim ii As Long
    Dim depthB As Long, depthP As Long, depthBr As Long
    For ii = 1 To stmtTokens.Count
        If stmtTokens(ii)(0) = "PAREN" Then
            Select Case stmtTokens(ii)(1)
                Case "{"
                    depthB = depthB + 1
                Case "}"
                    If depthB > 0 Then depthB = depthB - 1
                Case "("
                    depthP = depthP + 1
                Case ")"
                    If depthP > 0 Then depthP = depthP - 1
                Case "["
                    depthBr = depthBr + 1
                Case "]"
                    If depthBr > 0 Then depthBr = depthBr - 1
            End Select
        End If
        If stmtTokens(ii)(0) = "OP" And stmtTokens(ii)(1) = "=" And depthB = 0 And depthP = 0 And depthBr = 0 Then
            assignPos = ii
            Exit For
        End If
    Next ii

    If assignPos > 0 Then
        Dim lhsT As Collection, rhsT As Collection, t As Long
        Set lhsT = New Collection: Set rhsT = New Collection
        For t = 1 To assignPos - 1: lhsT.Add stmtTokens(t): Next t
        For t = assignPos + 1 To stmtTokens.Count: rhsT.Add stmtTokens(t): Next t
        Dim lhsNode As Map
        If lhsT.Count = 1 And lhsT(1)(0) = "IDENT" Then
            Dim nm As String: nm = lhsT(1)(1)
            If InStr(nm, "[") > 0 Then
                Dim baseName As String: baseName = left$(nm, InStr(nm, "[") - 1)
                Dim idxStr As String: idxStr = Mid$(nm, InStr(nm, "[") + 1, Len(nm) - InStr(nm, "[") - 1)
                Dim idxNode As Map: Set idxNode = ParseExprFromStringToNode(idxStr)
                Set lhsNode = MakeNode("Index")
                Dim tmpVar As Map: Set tmpVar = MakeNode("Variable")
                tmpVar.SetValue "name", baseName
                lhsNode.SetValue "base", tmpVar
                lhsNode.SetValue "index", idxNode
            Else
                Set lhsNode = MakeNode("Variable")
                lhsNode.SetValue "name", nm
            End If
        Else
            Set lhsNode = ParseExprTokensToNode(lhsT)
        End If
        Dim rhsNode As Map: Set rhsNode = ParseExprTokensToNode(rhsT)
        Dim asn As Map: Set asn = MakeNode("Assign")
        asn.SetValue "left", lhsNode
        asn.SetValue "right", rhsNode
        Set ParseStatementTokensToAST = asn
        Exit Function
    End If

    ' fallback: expression statement
    Dim exprNode As Map
    Set exprNode = ParseExprTokensToNode(stmtTokens)
    Dim es As Map: Set es = MakeNode("ExprStmt")
    es.SetValue "expr", exprNode
    Set ParseStatementTokensToAST = es
End Function

' ---------------------
' Expression parsing -> Expr AST nodes
' ---------------------

Private Function ParseArgsTokensToExprNodes(argTokens As Collection) As Collection
    Dim out As New Collection
    If argTokens.Count = 0 Then
        Set ParseArgsTokensToExprNodes = out
        Exit Function
    End If
    Dim i As Long: i = 1
    Dim braceDepth As Long: braceDepth = 0
    Dim parenDepth As Long: parenDepth = 0
    Dim bracketDepth As Long: bracketDepth = 0
    Dim cur As New Collection
ParseArgsMain:
    Do While i <= argTokens.Count
        If argTokens(i)(0) = "PAREN" Then
            Select Case argTokens(i)(1)
                Case "{"
                    braceDepth = braceDepth + 1
                Case "}"
                    If braceDepth > 0 Then braceDepth = braceDepth - 1
                Case "("
                    parenDepth = parenDepth + 1
                Case ")"
                    If parenDepth > 0 Then parenDepth = parenDepth - 1
                Case "["
                    bracketDepth = bracketDepth + 1
                Case "]"
                    If bracketDepth > 0 Then bracketDepth = bracketDepth - 1
            End Select
        End If
        ' Argument/element separator: comma only (semicolon is NOT an argument separator)
        If argTokens(i)(0) = "SEP" And argTokens(i)(1) = "," And braceDepth = 0 And parenDepth = 0 And bracketDepth = 0 Then
            If cur.Count > 0 Then out.Add ParseExprTokensToNode(cur)
            Set cur = New Collection
            i = i + 1
            GoTo ParseArgsMain
        End If

        cur.Add argTokens(i)
        i = i + 1
    Loop
    If cur.Count > 0 Then out.Add ParseExprTokensToNode(cur)
    Set ParseArgsTokensToExprNodes = out
End Function

Private Function ParseExprFromStringToNode(exprStr As String) As Map
    Dim p As Parser
    Set p = New Parser
    Dim toks As Collection
    Set toks = p.Tokenize(exprStr)
    Set ParseExprFromStringToNode = ParseExprTokensToNode(toks)
End Function

Private Function ParseExprTokensToNode(toks As Collection) As Map
    Dim n As Long: n = toks.Count
    If n = 0 Then
        Set ParseExprTokensToNode = MakeNode("Literal"): ParseExprTokensToNode.SetValue "value", Empty
        Exit Function
    End If

    Dim types() As String, vals() As Variant, i As Long
    ReDim types(0 To n - 1)
    ReDim vals(0 To n - 1)
    For i = 1 To n
        types(i - 1) = toks(i)(0)
        vals(i - 1) = toks(i)(1)
    Next i

    Dim idx As Long: idx = 0
    Set ParseExprTokensToNode = ParseTernaryNode(types, vals, n, idx)
End Function

' Recursive descent building nodes

Private Function ParseTernaryNode(types() As String, vals() As Variant, n As Long, ByRef idx As Long) As Map
    ' parse cond ? trueExpr : falseExpr (right-assoc)
    Dim cond As Map
    Set cond = ParseLogicalOrNode(types, vals, n, idx)
    If idx < n Then
        If types(idx) = "OP" And vals(idx) = "?" Then
            idx = idx + 1
            Dim trueExpr As Map
            Set trueExpr = ParseTernaryNode(types, vals, n, idx) ' right-assoc: allow nested ternary
            ' expect ':'
            If idx < n And types(idx) = "OP" And vals(idx) = ":" Then
                idx = idx + 1
            Else
                err.Raise vbObjectError + 8001, "Compiler.ParseTernaryNode", "Expected ':' in ternary expression"
            End If
            Dim falseExpr As Map
            Set falseExpr = ParseTernaryNode(types, vals, n, idx)
            Dim tn As New Map
            tn.Add "type", "Ternary"
            tn.SetValue "cond", cond
            tn.SetValue "trueExpr", trueExpr
            tn.SetValue "falseExpr", falseExpr
            Set ParseTernaryNode = tn
            Exit Function
        End If
    End If
    Set ParseTernaryNode = cond
End Function

Private Function ParseLogicalOrNode(types() As String, vals() As Variant, n As Long, ByRef idx As Long) As Map
    Dim left_ As Map
    Dim right_ As Map
    Dim node As Map
    Dim op As String
    Set left_ = ParseLogicalAndNode(types, vals, n, idx)
    Do While idx < n
        If Not (types(idx) = "OP" And vals(idx) = "||") Then Exit Do
        op = vals(idx)
        idx = idx + 1
        Set right_ = ParseLogicalAndNode(types, vals, n, idx)
        Set node = MakeNode("Binary")
        node.SetValue "op", op
        node.SetValue "left", left_.Clone
        node.SetValue "right", right_.Clone
        Set left_ = node
    Loop
    Set ParseLogicalOrNode = left_.Clone
    Set left_ = Nothing
    Set right_ = Nothing
    Set node = Nothing
End Function

Private Function ParseLogicalAndNode(types() As String, vals() As Variant, n As Long, ByRef idx As Long) As Map
    Dim left_ As Map
    Dim right_ As Map
    Dim node As Map
    Set left_ = ParseEqualityNode(types, vals, n, idx)
    Do While idx < n
        If Not (types(idx) = "OP" And vals(idx) = "&&") Then Exit Do
        idx = idx + 1
        Set right_ = ParseEqualityNode(types, vals, n, idx)
        Set node = MakeNode("Binary")
        node.SetValue "op", "&&"
        node.SetValue "left", left_.Clone
        node.SetValue "right", right_.Clone
        Set left_ = node
    Loop
    Set ParseLogicalAndNode = left_.Clone
    Set left_ = Nothing
    Set right_ = Nothing
    Set node = Nothing
End Function

Private Function ParseEqualityNode(types() As String, vals() As Variant, n As Long, ByRef idx As Long) As Map
    Dim left_ As Map
    Dim right_ As Map
    Dim node As Map
    Dim op As String
    Set left_ = ParseRelationalNode(types, vals, n, idx)
    Do While idx < n
        If Not (types(idx) = "OP" And (vals(idx) = "==" Or vals(idx) = "=" Or vals(idx) = "!=")) Then Exit Do
        op = vals(idx)
        idx = idx + 1
        Set right_ = ParseRelationalNode(types, vals, n, idx)
        Set node = MakeNode("Binary")
        node.SetValue "op", op
        node.SetValue "left", left_.Clone
        node.SetValue "right", right_.Clone
        Set left_ = node
    Loop
    Set ParseEqualityNode = left_.Clone
    Set left_ = Nothing
    Set right_ = Nothing
    Set node = Nothing
End Function

Private Function ParseRelationalNode(types() As String, vals() As Variant, n As Long, ByRef idx As Long) As Map
    Dim left_ As Map
    Dim right_ As Map
    Dim node As Map
    Dim op As String
    Set left_ = ParseAddNode(types, vals, n, idx)
    Do While idx < n
        If Not (types(idx) = "OP" And (vals(idx) = "<" Or vals(idx) = ">" Or vals(idx) = "<=" Or vals(idx) = ">=")) Then Exit Do
        op = vals(idx)
        idx = idx + 1
        Set right_ = ParseAddNode(types, vals, n, idx)
        Set node = MakeNode("Binary")
        node.SetValue "op", op
        node.SetValue "left", left_.Clone
        node.SetValue "right", right_.Clone
        Set left_ = node
    Loop
    Set ParseRelationalNode = left_.Clone
    Set left_ = Nothing
    Set right_ = Nothing
    Set node = Nothing
End Function

Private Function ParseAddNode(types() As String, vals() As Variant, n As Long, ByRef idx As Long) As Map
    Dim left_ As Map
    Dim op As String
    Dim right_ As Map
    Dim node As Map
    Set left_ = ParseMulNode(types, vals, n, idx)
    Do While idx < n
        If Not (types(idx) = "OP" And (vals(idx) = "+" Or vals(idx) = "-")) Then Exit Do
        op = vals(idx)
        idx = idx + 1
        Set right_ = ParseMulNode(types, vals, n, idx)
        Set node = MakeNode("Binary")
        node.SetValue "op", op
        node.SetValue "left", left_.Clone
        node.SetValue "right", right_.Clone
        Set left_ = node
    Loop
    Set ParseAddNode = left_.Clone
    Set left_ = Nothing
    Set right_ = Nothing
    Set node = Nothing
End Function

Private Function ParseMulNode(types() As String, vals() As Variant, n As Long, ByRef idx As Long) As Map
    Dim left_ As Map
    Dim right_ As Map
    Dim node As Map
    Dim op As String
    Set left_ = ParsePowNode(types, vals, n, idx)
    Do While idx < n
        If Not (types(idx) = "OP" And (vals(idx) = "*" Or vals(idx) = "/" Or vals(idx) = "%")) Then Exit Do
        op = vals(idx)
        idx = idx + 1
        Set right_ = ParsePowNode(types, vals, n, idx)
        Set node = MakeNode("Binary")
        node.SetValue "op", op
        node.SetValue "left", left_.Clone
        node.SetValue "right", right_.Clone
        Set left_ = node
    Loop
    Set ParseMulNode = left_.Clone
    Set left_ = Nothing
    Set right_ = Nothing
    Set node = Nothing
End Function

Private Function ParsePowNode(types() As String, vals() As Variant, n As Long, ByRef idx As Long) As Map
    Dim left_ As Map
    Dim right_ As Map
    Dim node As Map
    Set left_ = ParseUnaryNode(types, vals, n, idx)
    Do While idx < n
        ' Right-associative exponentiation:
        ' if we see '^' after the left operand, parse the RHS with ParsePow
        ' so a ^ b ^ c => a ^ (b ^ c)
        If Not (types(idx) = "OP" And vals(idx) = "^") Then Exit Do
        idx = idx + 1
        ' parse the right-hand side with ParsePow to ensure right-assoc
        Set right_ = ParsePowNode(types, vals, n, idx)
        Set node = MakeNode("Binary")
        node.SetValue "op", "^"
        node.SetValue "left", left_.Clone
        node.SetValue "right", right_.Clone
        Set left_ = node
    Loop
    Set ParsePowNode = left_.Clone
    Set left_ = Nothing
    Set right_ = Nothing
    Set node = Nothing
End Function

Private Function ParseUnaryNode(types() As String, vals() As Variant, n As Long, ByRef idx As Long) As Map
    If idx < n And types(idx) = "OP" Then
        Dim op As String: op = vals(idx)
        If op = "!" Or op = "-" Then
            idx = idx + 1
            Dim v As Map: Set v = ParseUnaryNode(types, vals, n, idx)
            Dim node As Map: Set node = MakeNode("Unary")
            node.SetValue "op", op
            node.SetValue "expr", v
            Set ParseUnaryNode = node
            Exit Function
        End If
    End If
    Set ParseUnaryNode = ParsePrimaryNode(types, vals, n, idx)
End Function

Private Function ParsePrimaryNode(types() As String, vals() As Variant, n As Long, ByRef idx As Long) As Map
    If idx >= n Then
        Set ParsePrimaryNode = MakeNode("Literal")
        ParsePrimaryNode.SetValue "value", Empty
        Exit Function
    End If
    Dim t As String: t = types(idx)
    Dim v As Variant: v = vals(idx)

    Select Case t
        Case "NUMBER"
            Dim nNode As Map: Set nNode = MakeNode("Literal")
            nNode.SetValue "value", CDbl(v)
            idx = idx + 1
            Set ParsePrimaryNode = nNode
            Exit Function
        Case "STRING"
            Dim sNode As Map: Set sNode = MakeNode("Literal")
            sNode.SetValue "value", v
            idx = idx + 1
            Set ParsePrimaryNode = sNode
            Exit Function
        Case "VBEXPR"
            ' Special token produced by the parser for @(...) or direct VBAexpressions block.
            ' Create a VBExpr node with the raw expression string (to be evaluated by VBAexpressions at runtime).
            Dim vbNode As Map: Set vbNode = MakeNode("VBExpr")
            vbNode.SetValue "expr", CStr(v)
            idx = idx + 1
            Set ParsePrimaryNode = vbNode
            Exit Function
        Case "IDENT"
            Dim name As String: name = v
            Dim idxNode As Map
            Dim idxExpr As Map
            If LCase$(name) = "true" Then
                Dim bt As Map: Set bt = MakeNode("Literal")
                bt.SetValue "value", True
                idx = idx + 1
                Set ParsePrimaryNode = bt
                Exit Function
            ElseIf LCase$(name) = "false" Then
                Dim bf As Map: Set bf = MakeNode("Literal")
                bf.SetValue "value", False
                idx = idx + 1
                Set ParsePrimaryNode = bf
                Exit Function
            End If

            ' Expression-level anonymous function literal: fun (p1, p2) { ... }
            ' This reuses the same 'fun' token used for top-level function declarations but emits
            ' a FuncLiteral node usable inside expressions.
            If LCase$(name) = "fun" Then
                ' advance past 'fun'
                idx = idx + 1
                ' expect parameter list
                Dim params As Collection
                Set params = New Collection
                If idx < n And types(idx) = "PAREN" And vals(idx) = "(" Then
                    idx = idx + 1
                    Do While idx < n And Not (types(idx) = "PAREN" And vals(idx) = ")")
                        If types(idx) = "IDENT" Then
                            params.Add CStr(vals(idx))
                            idx = idx + 1
                            If idx < n And types(idx) = "SEP" And vals(idx) = "," Then idx = idx + 1
                        ElseIf types(idx) = "SEP" And vals(idx) = "," Then
                            idx = idx + 1
                        Else
                            ' skip unexpected tokens until ')'
                            idx = idx + 1
                        End If
                    Loop
                    If idx < n And types(idx) = "PAREN" And vals(idx) = ")" Then idx = idx + 1
                End If

                ' parse body { ... }
                If idx < n And types(idx) = "PAREN" And vals(idx) = "{" Then
                    idx = idx + 1
                    Dim depthBody As Long: depthBody = 0
                    Dim bodyTok As Collection
                    Set bodyTok = New Collection
                    Do While idx < n
                        If types(idx) = "PAREN" Then
                            If vals(idx) = "{" Then
                                depthBody = depthBody + 1
                            ElseIf vals(idx) = "}" Then
                                If depthBody = 0 Then
                                    Exit Do
                                Else
                                    depthBody = depthBody - 1
                                End If
                            End If
                        End If
                        bodyTok.Add Array(types(idx), vals(idx))
                        idx = idx + 1
                    Loop
                    ' consume closing '}'
                    If idx < n And types(idx) = "PAREN" And vals(idx) = "}" Then idx = idx + 1
                    Dim bodyStmts As Collection: Set bodyStmts = ParseTokensToAST(bodyTok)
                    Dim fnode As Map: Set fnode = MakeNode("FuncLiteral")
                    fnode.SetValue "params", params
                    fnode.SetValue "body", bodyStmts
                    Set ParsePrimaryNode = fnode.Clone
                    Exit Function
                Else
                    ' no body found -> emit empty function literal
                    Dim fnode2 As Map: Set fnode2 = MakeNode("FuncLiteral")
                    fnode2.SetValue "params", params
                    Dim emptyCol As New Collection
                    fnode2.SetValue "body", emptyCol
                    Set ParsePrimaryNode = fnode2
                    Exit Function
                End If
            End If

            Dim lenCall As Map
            Dim builtinLen As Map
            Dim al As Collection
            ' .length special-case on collapsed IDENT token (emit Call with callee)
            If right$(name, 7) = ".length" Then
                Dim arrName As String: arrName = left$(name, Len(name) - 7)
                Dim varN As Map: Set varN = MakeNode("Variable")
                varN.SetValue "name", arrName
                Set lenCall = MakeNode("Call")
                ' Emit a call to the builtin callee ".__len__" with the array variable as sole arg.
                Set builtinLen = MakeNode("Variable")
                builtinLen.SetValue "name", ".__len__"
                ' Emit a call to builtin callee ".__len__" with the array variable as sole arg.
                Set builtinLen = MakeNode("Variable")
                builtinLen.SetValue "name", ".__len__"
                lenCall.SetValue "callee", builtinLen
                Set al = New Collection
                al.Add varN
                lenCall.SetValue "args", al
                idx = idx + 1
                Set ParsePrimaryNode = lenCall
                Exit Function
            End If

            ' Handle collapsed identifiers with array/index and optional trailing ".length"
            ' Examples:
            '   a[3]
            '   a[3].length
            If InStr(name, "[") > 0 Then
                Dim base As String: base = left$(name, InStr(name, "[") - 1)
                Dim idxStr As String: idxStr = Mid$(name, InStr(name, "[") + 1, Len(name) - InStr(name, "[") - 1)
                Dim idxNodeFromStr As Map: Set idxNodeFromStr = ParseExprFromStringToNode(idxStr)
                Dim inNode As Map: Set inNode = MakeNode("Index")
                Dim vn As Map: Set vn = MakeNode("Variable")
                vn.SetValue "name", base
                inNode.SetValue "base", vn
                inNode.SetValue "index", idxNodeFromStr
                idx = idx + 1
                ' If next tokens are ". length" treat as builtin length call on the indexed result.
                If idx < n Then
                    If types(idx) = "SYM" And vals(idx) = "." Then
                        If (idx + 1) < n And types(idx + 1) = "IDENT" And LCase$(CStr(vals(idx + 1))) = "length" Then
                            ' consume '.' and 'length'
                            idx = idx + 2
                            Set lenCall = MakeNode("Call")
                            ' set both 'callee' and 'name' for maximum compatibility with VM variants
                            Set builtinLen = MakeNode("Variable")
                            builtinLen.SetValue "name", ".__len__"
                            lenCall.SetValue "callee", builtinLen
                            lenCall.SetValue "name", ".__len__"
                            Set al = New Collection
                            al.Add inNode
                            lenCall.SetValue "args", al
                            Set ParsePrimaryNode = lenCall
                            Exit Function
                        End If
                    End If
                End If
                Set ParsePrimaryNode = inNode
                Exit Function
            End If
            ' create initial variable node for an IDENT and then apply postfix operators:
            Dim currentNode As Map: Set currentNode = MakeNode("Variable")
            currentNode.SetValue "name", name
            idx = idx + 1

            ' Postfix loop: handle .prop, [...], func calls (...) as postfixes chaining onto currentNode
            Do While idx < n
                ' member access a.b
                If types(idx) = "OP" And vals(idx) = "." Then
                    idx = idx + 1
                    If idx < n And types(idx) = "IDENT" Then
                        Dim mem As Map: Set mem = MakeNode("Member")
                        mem.SetValue "base", currentNode
                        mem.SetValue "prop", CStr(vals(idx))
                        Set currentNode = mem.Clone
                        idx = idx + 1
                        ' continue loop
                    Else
                        ' invalid member access; stop postfixing
                        Exit Do
                    End If
                ' index access a[expr]
                ElseIf types(idx) = "PAREN" And vals(idx) = "[" Then
                    idx = idx + 1
                    Dim idxTok As Collection
                    Set idxTok = New Collection
                    Dim depthIdx As Long: depthIdx = 0
                    Do While idx < n
                        If types(idx) = "PAREN" Then
                            If vals(idx) = "[" Then
                                depthIdx = depthIdx + 1
                            ElseIf vals(idx) = "]" Then
                                If depthIdx = 0 Then
                                    Exit Do
                                Else
                                    depthIdx = depthIdx - 1
                                End If
                            End If
                        End If
                        idxTok.Add Array(types(idx), vals(idx))
                        idx = idx + 1
                    Loop
                    Dim idxExprNode As Map: Set idxExprNode = ParseExprTokensToNode(idxTok)
                    If idx < n And types(idx) = "PAREN" And vals(idx) = "]" Then idx = idx + 1
                    Dim indexNode As Map: Set indexNode = MakeNode("Index")
                    indexNode.SetValue "base", currentNode
                    indexNode.SetValue "index", idxExprNode
                    Set currentNode = indexNode.Clone
                ' function / method call: ( arglist )
                ElseIf types(idx) = "PAREN" And vals(idx) = "(" Then
                    ' parse args with nested depth counters (borrowed pattern from previous call parsing logic)
                    idx = idx + 1
                    Dim argNodes As Collection
                    Dim argTok As Collection
                    Dim argTokDepthParen As Long
                    Dim argTokDepthBr As Long
                    Dim argTokDepthB As Long
                    Set argNodes = New Collection
                    Do
                        Set argTok = New Collection
                        argTokDepthParen = 0
                        argTokDepthBr = 0
                        argTokDepthB = 0
                        Do While idx < n
                            If types(idx) = "PAREN" Then
                                Select Case CStr(vals(idx))
                                    Case "("
                                        argTokDepthParen = argTokDepthParen + 1
                                    Case ")"
                                        If argTokDepthParen = 0 Then Exit Do Else argTokDepthParen = argTokDepthParen - 1
                                    Case "["
                                        argTokDepthBr = argTokDepthBr + 1
                                    Case "]"
                                        If argTokDepthBr = 0 Then Exit Do Else argTokDepthBr = argTokDepthBr - 1
                                    Case "{"
                                        argTokDepthB = argTokDepthB + 1
                                    Case "}"
                                        If argTokDepthB = 0 Then Exit Do Else argTokDepthB = argTokDepthB - 1
                                End Select
                            ElseIf types(idx) = "SEP" And vals(idx) = "," And argTokDepthParen = 0 And argTokDepthBr = 0 And argTokDepthB = 0 Then
                                Exit Do
                            End If
                            argTok.Add Array(types(idx), vals(idx))
                            idx = idx + 1
                        Loop
                        If argTok.Count > 0 Then
                            argNodes.Add ParseExprTokensToNode(argTok)
                        Else
                            Dim litEmpty As Map: Set litEmpty = MakeNode("Literal")
                            litEmpty.SetValue "value", Empty
                            argNodes.Add litEmpty.Clone
                        End If
                        If idx < n Then
                            If types(idx) = "SEP" And vals(idx) = "," Then
                                idx = idx + 1
                                ' continue parsing next arg
                            ElseIf types(idx) = "PAREN" And vals(idx) = ")" Then
                                idx = idx + 1
                                Exit Do
                            Else
                                ' unexpected token -> stop args parsing
                                Exit Do
                            End If
                        Else
                            Exit Do
                        End If
                    Loop
                    Dim callNode As Map: Set callNode = MakeNode("Call")
                    callNode.SetValue "callee", currentNode
                    callNode.SetValue "args", argNodes
                    Set currentNode = callNode.Clone
                Else
                    Exit Do
                End If
            Loop
            Set ParsePrimaryNode = currentNode.Clone
            Exit Function
        Case "PAREN"
            If v = "(" Then
                idx = idx + 1
                Dim innerNode As Map: Set innerNode = ParseLogicalOrNode(types, vals, n, idx)
                If idx < n And types(idx) = "PAREN" And vals(idx) = ")" Then idx = idx + 1
                Set ParsePrimaryNode = innerNode.Clone
                Exit Function
            ElseIf v = "[" Then
                idx = idx + 1
                Dim arrList As Collection
                Dim elemTok As Collection
                Set arrList = New Collection
                Do While idx < n And Not (types(idx) = "PAREN" And vals(idx) = "]")
                    Set elemTok = New Collection
                    ' Collect element tokens until a comma (element separator) or closing bracket
                    Do While idx < n And Not ((types(idx) = "SEP" And vals(idx) = ",") Or (types(idx) = "PAREN" And vals(idx) = "]"))
                        elemTok.Add Array(types(idx), vals(idx))
                        idx = idx + 1
                    Loop
                    arrList.Add ParseExprTokensToNode(elemTok)
                    If idx < n And types(idx) = "SEP" And vals(idx) = "," Then idx = idx + 1
                Loop
                If idx < n And types(idx) = "PAREN" And vals(idx) = "]" Then idx = idx + 1
                Dim arrNode As Map: Set arrNode = MakeNode("Array")
                arrNode.SetValue "items", arrList
                Set ParsePrimaryNode = arrNode.Clone
                Exit Function
            ElseIf v = "{" Then
                ' object literal: { key: value, key2: value2 }
                idx = idx + 1
                Dim objItems As Collection
                Dim pair As Collection
                Set objItems = New Collection
                ' key can be IDENT or STRING
                Dim keyTok As String
                Do While idx < n And Not (types(idx) = "PAREN" And vals(idx) = "}")
                    If types(idx) = "IDENT" Or types(idx) = "STRING" Then
                        keyTok = CStr(vals(idx))
                        idx = idx + 1
                    Else
                        err.Raise vbObjectError + 8002, "Compiler.ParsePrimary", "Invalid object key"
                    End If
                    ' expect ':'
                    If idx < n And types(idx) = "OP" And vals(idx) = ":" Then
                        idx = idx + 1
                    Else
                        err.Raise vbObjectError + 8003, "Compiler.ParsePrimary", "Expected ':' after object key"
                    End If
                    ' parse value expression
                    Dim valNode As Map
                    Set valNode = ParseTernaryNode(types, vals, n, idx)
                    ' store pair (key, node)
                    Set pair = New Collection
                    pair.Add keyTok
                    pair.Add valNode.Clone
                    objItems.Add pair
                    ' optional comma
                    If idx < n And types(idx) = "SEP" And vals(idx) = "," Then idx = idx + 1
                Loop
                If idx < n And types(idx) = "PAREN" And vals(idx) = "}" Then idx = idx + 1
                Dim onode As Map: Set onode = MakeNode("Object")
                onode.SetValue "items", objItems
                Set ParsePrimaryNode = onode.Clone
                Exit Function
            End If
    End Select

    Dim defN As Map: Set defN = MakeNode("Literal")
    defN.SetValue "value", Empty
    Set ParsePrimaryNode = defN
End Function

' ---------------------
' AST builders for control structures
' ---------------------

Private Function ParseIfAST(stmtTokens As Collection) As Map
    Dim node As Map: Set node = MakeNode("If")
    Dim i As Long: i = 2
    Do While i <= stmtTokens.Count And Not (stmtTokens(i)(0) = "PAREN" And stmtTokens(i)(1) = "(")
        i = i + 1
    Loop
    If i > stmtTokens.Count Then Set ParseIfAST = node: Exit Function
    i = i + 1
    Dim condTokens As Collection
    Set condTokens = New Collection
    Dim condDepth As Long: condDepth = 0
    Do While i <= stmtTokens.Count
        If stmtTokens(i)(0) = "PAREN" Then
            If stmtTokens(i)(1) = "(" Then
                condDepth = condDepth + 1
            ElseIf stmtTokens(i)(1) = ")" Then
                If condDepth = 0 Then
                    Exit Do
                Else
                    condDepth = condDepth - 1
                End If
            End If
        End If
        condTokens.Add stmtTokens(i)
        i = i + 1
    Loop
    node.SetValue "cond", ParseExprTokensToNode(condTokens)

    Do While i <= stmtTokens.Count And Not (stmtTokens(i)(0) = "PAREN" And stmtTokens(i)(1) = "{")
        i = i + 1
    Loop
    If i > stmtTokens.Count Then Set ParseIfAST = node.Clone: GoTo clean_
    Dim j As Long: j = i + 1
    Dim depth As Long: depth = 0
    Dim thenTokens As Collection
    Set thenTokens = New Collection
    Do While j <= stmtTokens.Count
        If stmtTokens(j)(0) = "PAREN" Then
            If stmtTokens(j)(1) = "{" Then
                depth = depth + 1
            ElseIf stmtTokens(j)(1) = "}" Then
                If depth = 0 Then
                    Exit Do
                Else
                    depth = depth - 1
                End If
            End If
        End If
        thenTokens.Add stmtTokens(j)
        j = j + 1
    Loop
    node.SetValue "then", ParseTokensToAST(thenTokens)

    Dim pos As Long: pos = j + 1
    Dim elseifConds As Collection, elseifBlocks As Collection
    Dim hasElse As Boolean: hasElse = False
    Dim tname As String
    Dim eCondTok As Collection
    Dim eBlockTok As Collection
    Dim elseTok As Collection
    Dim eDepth As Long
    Set elseifConds = New Collection
    Set elseifBlocks = New Collection
    Do While pos <= stmtTokens.Count
        If stmtTokens(pos)(0) = "COMMENT" Then
            pos = pos + 1
            GoTo IfParseNext
        End If
        If stmtTokens(pos)(0) = "IDENT" Then
            tname = LCase$(stmtTokens(pos)(1))
            If tname = "elseif" Then
                pos = pos + 1
                Do While pos <= stmtTokens.Count And Not (stmtTokens(pos)(0) = "PAREN" And stmtTokens(pos)(1) = "(")
                    pos = pos + 1
                Loop
                If pos > stmtTokens.Count Then Exit Do
                pos = pos + 1
                eDepth = 0
                Set eCondTok = New Collection
                Do While pos <= stmtTokens.Count
                    If stmtTokens(pos)(0) = "PAREN" Then
                        If stmtTokens(pos)(1) = "(" Then
                            eDepth = eDepth + 1
                        ElseIf stmtTokens(pos)(1) = ")" Then
                            If eDepth = 0 Then
                                Exit Do
                            Else
                                eDepth = eDepth - 1
                            End If
                        End If
                    End If
                    eCondTok.Add stmtTokens(pos)
                    pos = pos + 1
                Loop
                Do While pos <= stmtTokens.Count And Not (stmtTokens(pos)(0) = "PAREN" And stmtTokens(pos)(1) = "{")
                    pos = pos + 1
                Loop
                If pos > stmtTokens.Count Then Exit Do
                pos = pos + 1
                Set eBlockTok = New Collection
                depth = 0
                Do While pos <= stmtTokens.Count
                    If stmtTokens(pos)(0) = "PAREN" Then
                        If stmtTokens(pos)(1) = "{" Then
                            depth = depth + 1
                        ElseIf stmtTokens(pos)(1) = "}" Then
                            If depth = 0 Then
                                Exit Do
                            Else
                                depth = depth - 1
                            End If
                        End If
                    End If
                    eBlockTok.Add stmtTokens(pos)
                    pos = pos + 1
                Loop
                elseifConds.Add ParseExprTokensToNode(eCondTok)
                elseifBlocks.Add ParseTokensToAST(eBlockTok)
                pos = pos + 1
                GoTo IfParseNext
            ElseIf tname = "else" Then
                pos = pos + 1
                Do While pos <= stmtTokens.Count And Not (stmtTokens(pos)(0) = "PAREN" And stmtTokens(pos)(1) = "{")
                    pos = pos + 1
                Loop
                If pos > stmtTokens.Count Then Exit Do
                pos = pos + 1
                Set elseTok = New Collection
                depth = 0
                Do While pos <= stmtTokens.Count
                    If stmtTokens(pos)(0) = "PAREN" Then
                        If stmtTokens(pos)(1) = "{" Then
                            depth = depth + 1
                        ElseIf stmtTokens(pos)(1) = "}" Then
                            If depth = 0 Then
                                Exit Do
                            Else
                                depth = depth - 1
                            End If
                        End If
                    End If
                    elseTok.Add stmtTokens(pos)
                    pos = pos + 1
                Loop
                hasElse = True
                node.SetValue "else", ParseTokensToAST(elseTok)
                Exit Do
            Else
                Exit Do
            End If
        Else
            Exit Do
        End If
IfParseNext:
    Loop

    node.SetValue "elseif_conds", CloneCollectionOfVariants(elseifConds)
    node.SetValue "elseif_blocks", CloneCollectionOfVariants(elseifBlocks)
    node.SetValue "hasElse", hasElse
    Set ParseIfAST = node
clean_:
    Set node = Nothing
    Set eCondTok = Nothing
    Set eBlockTok = Nothing
    Set elseTok = Nothing
    Set elseifConds = Nothing
    Set elseifBlocks = Nothing
End Function

Private Function ParseForAST(stmtTokens As Collection) As Map
    Dim node As Map: Set node = MakeNode("For")
    Dim initTok As Collection
    Dim condTok As Collection
    Dim stepTok As Collection
    Dim bodyToks As Collection
    Dim i As Long: i = 2
    Do While i <= stmtTokens.Count And Not (stmtTokens(i)(0) = "PAREN" And stmtTokens(i)(1) = "(")
        i = i + 1
    Loop
    If i > stmtTokens.Count Then Set ParseForAST = node: Exit Function
    i = i + 1
    Set initTok = New Collection
    Dim hdrParen As Long, hdrBracket As Long, hdrBrace As Long
    hdrParen = 0: hdrBracket = 0: hdrBrace = 0
    Do While i <= stmtTokens.Count
        If stmtTokens(i)(0) = "PAREN" Then
            Select Case stmtTokens(i)(1)
                Case "("
                    hdrParen = hdrParen + 1
                Case ")"
                    If hdrParen > 0 Then
                        hdrParen = hdrParen - 1
                    Else
                        Exit Do
                    End If
                Case "["
                    hdrBracket = hdrBracket + 1
                Case "]"
                    If hdrBracket > 0 Then hdrBracket = hdrBracket - 1
                Case "{"
                    hdrBrace = hdrBrace + 1
                Case "}"
                    If hdrBrace > 0 Then hdrBrace = hdrBrace - 1
            End Select
        End If
        If hdrParen = 0 And hdrBracket = 0 And hdrBrace = 0 And stmtTokens(i)(0) = "SEP" And _
                        (stmtTokens(i)(1) = "," Or stmtTokens(i)(1) = ";") Then
            i = i + 1
            Exit Do
        End If
        initTok.Add stmtTokens(i)
        i = i + 1
    Loop
    node.SetValue "init", ParseExprTokensToNode(initTok)

    Set condTok = New Collection
    hdrParen = 0: hdrBracket = 0: hdrBrace = 0
    Do While i <= stmtTokens.Count
        If stmtTokens(i)(0) = "PAREN" Then
            Select Case stmtTokens(i)(1)
                Case "("
                    hdrParen = hdrParen + 1
                Case ")"
                    If hdrParen > 0 Then
                        hdrParen = hdrParen - 1
                    Else
                        Exit Do
                    End If
                Case "["
                    hdrBracket = hdrBracket + 1
                Case "]"
                    If hdrBracket > 0 Then hdrBracket = hdrBracket - 1
                Case "{"
                    hdrBrace = hdrBrace + 1
                Case "}"
                    If hdrBrace > 0 Then hdrBrace = hdrBrace - 1
            End Select
        End If
        If hdrParen = 0 And hdrBracket = 0 And hdrBrace = 0 And stmtTokens(i)(0) = "SEP" And _
                        (stmtTokens(i)(1) = "," Or stmtTokens(i)(1) = ";") Then
            i = i + 1
            Exit Do
        End If
        condTok.Add stmtTokens(i)
        i = i + 1
    Loop
    node.SetValue "cond", ParseExprTokensToNode(condTok)

    Set stepTok = New Collection
    hdrParen = 0: hdrBracket = 0: hdrBrace = 0
    Do While i <= stmtTokens.Count And Not (stmtTokens(i)(0) = "PAREN" And stmtTokens(i)(1) = ")")
        If stmtTokens(i)(0) = "PAREN" Then
            Select Case stmtTokens(i)(1)
                Case "("
                    hdrParen = hdrParen + 1
                Case ")"
                    If hdrParen > 0 Then hdrParen = hdrParen - 1
                Case "["
                    hdrBracket = hdrBracket + 1
                Case "]"
                    If hdrBracket > 0 Then hdrBracket = hdrBracket - 1
                Case "{"
                    hdrBrace = hdrBrace + 1
                Case "}"
                    If hdrBrace > 0 Then hdrBrace = hdrBrace - 1
            End Select
        End If
        stepTok.Add stmtTokens(i)
        i = i + 1
    Loop
    node.SetValue "step", ParseExprTokensToNode(stepTok)

    Do While i <= stmtTokens.Count And Not (stmtTokens(i)(0) = "PAREN" And stmtTokens(i)(1) = "{")
        i = i + 1
    Loop
    If i > stmtTokens.Count Then Set ParseForAST = node.Clone: GoTo clean_
    Dim j As Long: j = i + 1
    Dim depth As Long: depth = 0
    Set bodyToks = New Collection
    Do While j <= stmtTokens.Count
        If stmtTokens(j)(0) = "PAREN" Then
            If stmtTokens(j)(1) = "{" Then
                depth = depth + 1
            ElseIf stmtTokens(j)(1) = "}" Then
                If depth = 0 Then
                    Exit Do
                Else
                    depth = depth - 1
                End If
            End If
        End If
        bodyToks.Add stmtTokens(j)
        j = j + 1
    Loop
    node.SetValue "body", ParseTokensToAST(bodyToks)
    Set ParseForAST = node.Clone
clean_:
    Set node = Nothing
    Set bodyToks = Nothing
    Set stepTok = Nothing
    Set condTok = Nothing
    Set initTok = Nothing
End Function

Private Function ParseWhileAST(stmtTokens As Collection) As Map
    Dim node As Map: Set node = MakeNode("While")
    Dim i As Long: i = 2
    Do While i <= stmtTokens.Count And Not (stmtTokens(i)(0) = "PAREN" And stmtTokens(i)(1) = "(")
        i = i + 1
    Loop
    If i > stmtTokens.Count Then Set ParseWhileAST = node: Exit Function
    i = i + 1
    Dim condTok As Collection
    Set condTok = New Collection
    Dim depthP As Long: depthP = 0
    Do While i <= stmtTokens.Count
        If stmtTokens(i)(0) = "PAREN" Then
            If stmtTokens(i)(1) = "(" Then
                depthP = depthP + 1
            ElseIf stmtTokens(i)(1) = ")" Then
                If depthP = 0 Then
                    Exit Do
                Else
                    depthP = depthP - 1
                End If
            End If
        End If
        condTok.Add stmtTokens(i)
        i = i + 1
    Loop
    node.SetValue "cond", ParseExprTokensToNode(condTok)

    Do While i <= stmtTokens.Count And Not (stmtTokens(i)(0) = "PAREN" And stmtTokens(i)(1) = "{")
        i = i + 1
    Loop
    If i > stmtTokens.Count Then Set ParseWhileAST = node: Exit Function
    Dim j As Long: j = i + 1
    Dim depth As Long: depth = 0
    Dim bodyToks As Collection
    Set bodyToks = New Collection
    Do While j <= stmtTokens.Count
        If stmtTokens(j)(0) = "PAREN" Then
            If stmtTokens(j)(1) = "{" Then
                depth = depth + 1
            ElseIf stmtTokens(j)(1) = "}" Then
                If depth = 0 Then
                    Exit Do
                Else
                    depth = depth - 1
                End If
            End If
        End If
        bodyToks.Add stmtTokens(j)
        j = j + 1
    Loop
    node.SetValue "body", ParseTokensToAST(bodyToks)
    Set ParseWhileAST = node
End Function

Private Function ParseTryCatchAST(stmtTokens As Collection) As Map
    Dim node As Map: Set node = MakeNode("TryCatch")
    Dim i As Long: i = 2
    Do While i <= stmtTokens.Count And Not (stmtTokens(i)(0) = "PAREN" And stmtTokens(i)(1) = "{")
        i = i + 1
    Loop
    If i > stmtTokens.Count Then Set ParseTryCatchAST = node: Exit Function
    Dim j As Long: j = i + 1
    Dim depth As Long: depth = 0
    Dim tryTok As Collection
    Set tryTok = New Collection
    Do While j <= stmtTokens.Count
        If stmtTokens(j)(0) = "PAREN" Then
            If stmtTokens(j)(1) = "{" Then
                depth = depth + 1
            ElseIf stmtTokens(j)(1) = "}" Then
                If depth = 0 Then
                    Exit Do
                Else
                    depth = depth - 1
                End If
            End If
        End If
        tryTok.Add stmtTokens(j)
        j = j + 1
    Loop
    node.SetValue "try", ParseTokensToAST(tryTok)

    Dim k As Long: k = j + 1
    Dim catchTok As Collection
    Set catchTok = New Collection
    Do While k <= stmtTokens.Count
        If stmtTokens(k)(0) = "IDENT" And LCase$(stmtTokens(k)(1)) = "catch" Then
            Dim kk As Long: kk = k + 1
            Do While kk <= stmtTokens.Count And Not (stmtTokens(kk)(0) = "PAREN" And stmtTokens(kk)(1) = "{")
                kk = kk + 1
            Loop
            If kk <= stmtTokens.Count Then
                kk = kk + 1
                Dim depth2 As Long: depth2 = 0
                Do While kk <= stmtTokens.Count
                    If stmtTokens(kk)(0) = "PAREN" Then
                        If stmtTokens(kk)(1) = "{" Then
                            depth2 = depth2 + 1
                        ElseIf stmtTokens(kk)(1) = "}" Then
                            If depth2 = 0 Then
                                Exit Do
                            Else
                                depth2 = depth2 - 1
                            End If
                        End If
                    End If
                    catchTok.Add stmtTokens(kk)
                    kk = kk + 1
                Loop
            End If
            Exit Do
        End If
        k = k + 1
    Loop
    node.SetValue "catch", ParseTokensToAST(catchTok)
    Set ParseTryCatchAST = node
End Function

Private Function ParseSwitchAST(stmtTokens As Collection) As Map
    Dim node As Map: Set node = MakeNode("Switch")
    Dim i As Long: i = 2
    Do While i <= stmtTokens.Count And Not (stmtTokens(i)(0) = "PAREN" And stmtTokens(i)(1) = "(")
        i = i + 1
    Loop
    If i > stmtTokens.Count Then Set ParseSwitchAST = node: Exit Function
    i = i + 1
    Dim condTok As New Collection
    Do While i <= stmtTokens.Count And Not (stmtTokens(i)(0) = "PAREN" And stmtTokens(i)(1) = ")")
        condTok.Add stmtTokens(i)
        i = i + 1
    Loop
    node.SetValue "expr", ParseExprTokensToNode(condTok)

    Do While i <= stmtTokens.Count And Not (stmtTokens(i)(0) = "PAREN" And stmtTokens(i)(1) = "{")
        i = i + 1
    Loop
    If i > stmtTokens.Count Then Set ParseSwitchAST = node: Exit Function
    i = i + 1
    Dim cases As Collection
    Set cases = New Collection
    Dim defaultBlock As Collection
    Dim valTok As Collection
    Do While i <= stmtTokens.Count And Not (stmtTokens(i)(0) = "PAREN" And stmtTokens(i)(1) = "}")
        If stmtTokens(i)(0) = "IDENT" And LCase$(stmtTokens(i)(1)) = "case" Then
            i = i + 1
            Set valTok = New Collection
            Dim cvParen As Long: cvParen = 0
            Dim cvBracket As Long: cvBracket = 0
            Dim cvBrace As Long: cvBrace = 0
            Do While i <= stmtTokens.Count
                If stmtTokens(i)(0) = "PAREN" Then
                    Select Case stmtTokens(i)(1)
                        Case "("
                            cvParen = cvParen + 1
                        Case ")"
                            If cvParen > 0 Then cvParen = cvParen - 1
                        Case "["
                            cvBracket = cvBracket + 1
                        Case "]"
                            If cvBracket > 0 Then cvBracket = cvBracket - 1
                        Case "{"
                            If cvParen = 0 And cvBracket = 0 Then
                                Exit Do
                            Else
                                cvBrace = cvBrace + 1
                            End If
                        Case "}"
                            If cvBrace > 0 Then cvBrace = cvBrace - 1
                    End Select
                End If
                valTok.Add stmtTokens(i)
                i = i + 1
            Loop
            Dim caseValNode As Map: Set caseValNode = ParseExprTokensToNode(valTok)
            If i <= stmtTokens.Count And stmtTokens(i)(0) = "PAREN" And stmtTokens(i)(1) = "{" Then
                i = i + 1
                Dim depth As Long: depth = 0
                Dim blockTok As Collection
                Dim pair As Collection
                Dim caseBlock As Collection
                Set blockTok = New Collection
                Do While i <= stmtTokens.Count
                    If stmtTokens(i)(0) = "PAREN" Then
                        If stmtTokens(i)(1) = "{" Then
                            depth = depth + 1
                        ElseIf stmtTokens(i)(1) = "}" Then
                            If depth = 0 Then
                                Exit Do
                            Else
                                depth = depth - 1
                            End If
                        End If
                    End If
                    blockTok.Add stmtTokens(i)
                    i = i + 1
                Loop
                Set caseBlock = ParseTokensToAST(blockTok)
                Set pair = New Collection
                pair.Add caseValNode
                pair.Add caseBlock
                cases.Add pair
                i = i + 1
            End If
        ElseIf stmtTokens(i)(0) = "IDENT" And LCase$(stmtTokens(i)(1)) = "default" Then
            i = i + 1
            Do While i <= stmtTokens.Count And Not (stmtTokens(i)(0) = "PAREN" And stmtTokens(i)(1) = "{")
                i = i + 1
            Loop
            If i <= stmtTokens.Count And stmtTokens(i)(0) = "PAREN" And stmtTokens(i)(1) = "{" Then
                i = i + 1
                Dim depth2 As Long: depth2 = 0
                Dim defTok As New Collection
                Do While i <= stmtTokens.Count
                    If stmtTokens(i)(0) = "PAREN" Then
                        If stmtTokens(i)(1) = "{" Then
                            depth2 = depth2 + 1
                        ElseIf stmtTokens(i)(1) = "}" Then
                            If depth2 = 0 Then
                                Exit Do
                            Else
                                depth2 = depth2 - 1
                            End If
                        End If
                    End If
                    defTok.Add stmtTokens(i)
                    i = i + 1
                Loop
                Set defaultBlock = ParseTokensToAST(defTok)
                i = i + 1
            End If
        Else
            i = i + 1
        End If
    Loop
    node.SetValue "cases", cases
    node.SetValue "default", defaultBlock
    Set ParseSwitchAST = node
End Function


' Compiler AST normalization pass
' Converts expression statements like (Binary op "=" ...) into Assign nodes
' when the left-hand side is a valid assignment target (Variable or Index).
' Also normalizes `for` init/step and recursively visits blocks.
'
Private Function IsAssignableNode(n As Map) As Boolean
    If n Is Nothing Then
        IsAssignableNode = False: Exit Function
    End If
    Dim t As String: t = n.GetValue("type")
    ' Accept variable, indexed access (arr[i]) and member access (obj.prop)
    If t = "Variable" Or t = "Index" Or t = "Member" Then
        IsAssignableNode = True
    Else
        IsAssignableNode = False
    End If
End Function

Private Function NormalizeExprNodeRecursive(expr As Map) As Map
    ' Recursively walk expression nodes and normalize inner expressions as needed.
    If expr Is Nothing Then
        Set NormalizeExprNodeRecursive = Nothing: Exit Function
    End If
    Dim t As String: t = expr.GetValue("type")
    Select Case t
        Case "Binary"
            Dim left_ As Map: Set left_ = NormalizeExprNodeRecursive(expr.GetValue("left"))
            Dim right_ As Map: Set right_ = NormalizeExprNodeRecursive(expr.GetValue("right"))
            ' create new Binary node only if children changed, otherwise return original
            If (Not left_ Is expr.GetValue("left")) Or (Not right_ Is expr.GetValue("right")) Then
                Dim nb As New Map
                nb.Add "type", "Binary"
                nb.SetValue "op", expr.GetValue("op")
                nb.SetValue "left", IIf(left_ Is Nothing, expr.GetValue("left"), left_)
                nb.SetValue "right", IIf(right_ Is Nothing, expr.GetValue("right"), right_)
                Set NormalizeExprNodeRecursive = nb
                Exit Function
            End If
            Set NormalizeExprNodeRecursive = expr
            Exit Function
        Case "Unary"
            Dim rec As Map: Set rec = NormalizeExprNodeRecursive(expr.GetValue("expr"))
            If Not rec Is expr.GetValue("expr") Then
                Dim nu As New Map
                nu.Add "type", "Unary"
                nu.SetValue "op", expr.GetValue("op")
                nu.SetValue "expr", rec
                Set NormalizeExprNodeRecursive = nu
                Exit Function
            End If
            Set NormalizeExprNodeRecursive = expr
            Exit Function
        Case "Call"
            ' normalize args
            Dim args As Collection: Set args = expr.GetValue("args")
            If Not args Is Nothing Then
                Dim newArgs As New Collection
                Dim changed As Boolean: changed = False
                Dim i As Long
                For i = 1 To args.Count
                    Dim ae As Map: Set ae = NormalizeExprNodeRecursive(args(i))
                    If Not ae Is args(i) Then changed = True: newArgs.Add ae Else newArgs.Add args(i)
                Next i
                If changed Then
                    Dim nc As New Map
                    nc.Add "type", "Call"
                    nc.SetValue "name", expr.GetValue("name")
                    nc.SetValue "args", newArgs
                    Set NormalizeExprNodeRecursive = nc
                    Exit Function
                End If
            End If
            Set NormalizeExprNodeRecursive = expr
            Exit Function
        Case "Index"
            Dim b As Map: Set b = NormalizeExprNodeRecursive(expr.GetValue("base"))
            Dim idx As Map: Set idx = NormalizeExprNodeRecursive(expr.GetValue("index"))
            If Not b Is expr.GetValue("base") Or Not idx Is expr.GetValue("index") Then
                Dim ni As New Map
                ni.Add "type", "Index"
                ni.SetValue "base", IIf(b Is Nothing, expr.GetValue("base"), b)
                ni.SetValue "index", IIf(idx Is Nothing, expr.GetValue("index"), idx)
                Set NormalizeExprNodeRecursive = ni
                Exit Function
            End If
            Set NormalizeExprNodeRecursive = expr
            Exit Function
        Case "Array"
            Dim items As Collection: Set items = expr.GetValue("items")
            If Not items Is Nothing Then
                Dim newit As New Collection
                Dim ch As Boolean: ch = False
                Dim ii As Long
                For ii = 1 To items.Count
                    Dim ei As Map: Set ei = NormalizeExprNodeRecursive(items(ii))
                    If Not ei Is items(ii) Then ch = True: newit.Add ei Else newit.Add items(ii)
                Next ii
                If ch Then
                    Dim na As New Map
                    na.Add "type", "Array"
                    na.SetValue "items", newit
                    Set NormalizeExprNodeRecursive = na
                    Exit Function
                End If
            End If
            Set NormalizeExprNodeRecursive = expr
            Exit Function
        Case Else
            Set NormalizeExprNodeRecursive = expr
            Exit Function
    End Select
End Function

Private Function NormalizeNodeRecursive(node As Map) As Map
    ' Normalize a statement node and recursively process inner blocks.
    If node Is Nothing Then
        Set NormalizeNodeRecursive = Nothing: Exit Function
    End If
    Dim expr As Map
    Dim left_ As Map
    Dim a As Map
    Dim r As Map
    Dim newExpr As Map
    Dim ns As Map
    Dim cond As Map
    Dim thenBlk As Collection
    Dim outThen As Collection
    Dim elseifConds As Collection
    Dim elseifBlocks As Collection
    Dim outElseIfConds As Collection
    Dim outElseIfBlocks As Collection
    Dim eb As Collection
    Dim nb As Collection
    Dim elseBlk As Collection
    Dim outElse As Collection
    Dim nif As Map
    Dim initNode As Map
    Dim condNode As Map
    Dim stepNode As Map
    Dim body As Collection
    Dim nInit As Map, nCond As Map, nStep As Map
    Dim na As Map
    Dim L As Map
    Dim e As Map
    Dim L2 As Map
    Dim na2 As Map
    Dim SL As Map
    Dim na3 As Map
    Dim newBody As Collection
    Dim nf As Map
    Dim cnd As Map
    Dim bdy As Collection
    Dim nbdy As Collection
    Dim nw As Map
    Dim tryBlk As Collection
    Dim catchBlk As Collection
    Dim nt As Collection, nc As Collection
    Dim ntc As Map
    Dim cases As Collection
    Dim ncases As Collection
    Dim pair As Collection
    Dim caseExpr As Map
    Dim blockStmts As Collection
    Dim newPair As Collection
    Dim def As Collection
    Dim ndef As Collection
    Dim nsw As Map
    Dim leftA As Map
    Dim rightA As Map
    Dim naNode As Map
    Dim tp As String: tp = node.GetValue("type")
    Select Case tp
        Case "ExprStmt"
            Set expr = node.GetValue("expr")
            ' If expression is a Binary "=" with assignable LHS, convert to Assign node.
            If Not expr Is Nothing Then
                If expr.GetValue("type") = "Binary" Then
                    If CStr(expr.GetValue("op")) = "=" Then
                        Set left_ = expr.GetValue("left")
                        If Not left_ Is Nothing Then
                            If IsAssignableNode(left_) Then
                                a.Add "type", "Assign"
                                a.SetValue "left", left_
                                ' right: normalize recursively as expression
                                Set r = NormalizeExprNodeRecursive(expr.GetValue("right"))
                                If r Is Nothing Then Set r = expr.GetValue("right")
                                a.SetValue "right", r
                                Set NormalizeNodeRecursive = a
                                Exit Function
                            End If
                        End If
                    End If
                End If
            End If
            ' otherwise just normalize inner expression (if any)
            If Not expr Is Nothing Then
                Set newExpr = NormalizeExprNodeRecursive(expr)
                If Not newExpr Is expr Then
                    Set ns = New Map
                    ns.Add "type", "ExprStmt"
                    ns.SetValue "expr", newExpr
                    Set NormalizeNodeRecursive = ns
                    Exit Function
                End If
            End If
            Set NormalizeNodeRecursive = node
            Exit Function

        Case "If"
            Set cond = NormalizeExprNodeRecursive(node.GetValue("cond"))
            Set thenBlk = node.GetValue("then")
            Dim i As Long
            If Not thenBlk Is Nothing Then
                Set outThen = New Collection
                For i = 1 To thenBlk.Count
                    outThen.Add NormalizeNodeRecursive(thenBlk(i))
                Next i
            End If
            Set elseifConds = node.GetValue("elseif_conds")
            Set elseifBlocks = node.GetValue("elseif_blocks")
            Set outElseIfConds = New Collection
            Set outElseIfBlocks = New Collection
            If Not elseifConds Is Nothing Then
                For i = 1 To elseifConds.Count
                    outElseIfConds.Add NormalizeExprNodeRecursive(elseifConds(i))
                Next i
            End If
            If Not elseifBlocks Is Nothing Then
                For i = 1 To elseifBlocks.Count
                    Set eb = elseifBlocks(i)
                    Set nb = New Collection
                    Dim j As Long
                    For j = 1 To eb.Count
                        nb.Add NormalizeNodeRecursive(eb(j))
                    Next j
                    outElseIfBlocks.Add nb
                Next i
            End If
            Set elseBlk = node.GetValue("else")
            If Not elseBlk Is Nothing Then
                Set outElse = New Collection
                For i = 1 To elseBlk.Count
                    outElse.Add NormalizeNodeRecursive(elseBlk(i))
                Next i
            End If
            Set nif = New Map
            nif.Add "type", "If"
            nif.SetValue "cond", IIf(cond Is Nothing, node.GetValue("cond"), cond)
            nif.SetValue "then", outThen
            nif.SetValue "elseif_conds", outElseIfConds
            nif.SetValue "elseif_blocks", outElseIfBlocks
            nif.SetValue "hasElse", node.GetValue("hasElse")
            If Not outElse Is Nothing Then nif.SetValue "else", outElse
            Set NormalizeNodeRecursive = nif
            Exit Function

        Case "For"
            ' normalize init/cond/step and body
            Set initNode = node.GetValue("init")
            Set condNode = node.GetValue("cond")
            Set stepNode = node.GetValue("step")
            Set body = node.GetValue("body")
            If Not initNode Is Nothing Then
                If initNode.GetValue("type") = "Binary" And CStr(initNode.GetValue("op")) = "=" Then
                    Set L = initNode.GetValue("left")
                    If Not L Is Nothing And IsAssignableNode(L) Then
                        Set na = New Map
                        na.Add "type", "Assign"
                        na.SetValue "left", L
                        na.SetValue "right", NormalizeExprNodeRecursive(initNode.GetValue("right"))
                        Set nInit = na
                    Else
                        Set nInit = NormalizeExprNodeRecursive(initNode)
                    End If
                ElseIf initNode.GetValue("type") = "ExprStmt" Then
                    ' exprstmt could wrap a binary expression
                    Set e = initNode.GetValue("expr")
                    If Not e Is Nothing And e.GetValue("type") = "Binary" And CStr(e.GetValue("op")) = "=" Then
                        Set L2 = e.GetValue("left")
                        If Not L2 Is Nothing And IsAssignableNode(L2) Then
                            Set na2 = New Map
                            na2.Add "type", "Assign"
                            na2.SetValue "left", L2
                            na2.SetValue "right", NormalizeExprNodeRecursive(e.GetValue("right"))
                            Set nInit = na2
                        Else
                            Set nInit = NormalizeExprNodeRecursive(e)
                        End If
                    Else
                        Set nInit = NormalizeExprNodeRecursive(e)
                    End If
                Else
                    Set nInit = NormalizeExprNodeRecursive(initNode)
                End If
            End If
            If Not condNode Is Nothing Then Set nCond = NormalizeExprNodeRecursive(condNode)
            If Not stepNode Is Nothing Then
                If stepNode.GetValue("type") = "Binary" And CStr(stepNode.GetValue("op")) = "=" Then
                    Set SL = stepNode.GetValue("left")
                    If Not SL Is Nothing And IsAssignableNode(SL) Then
                        Set na3 = New Map
                        na3.Add "type", "Assign"
                        na3.SetValue "left", SL
                        na3.SetValue "right", NormalizeExprNodeRecursive(stepNode.GetValue("right"))
                        Set nStep = na3
                    Else
                        Set nStep = NormalizeExprNodeRecursive(stepNode)
                    End If
                Else
                    Set nStep = NormalizeExprNodeRecursive(stepNode)
                End If
            End If
            Set newBody = New Collection
            If Not body Is Nothing Then
                Dim bidx As Long
                For bidx = 1 To body.Count
                    newBody.Add NormalizeNodeRecursive(body(bidx))
                Next bidx
            End If
            Set nf = New Map
            nf.Add "type", "For"
            If Not nInit Is Nothing Then nf.SetValue "init", nInit
            If Not nCond Is Nothing Then nf.SetValue "cond", nCond
            If Not nStep Is Nothing Then nf.SetValue "step", nStep
            nf.SetValue "body", newBody
            Set NormalizeNodeRecursive = nf
            Exit Function

        Case "While"
            Set cnd = NormalizeExprNodeRecursive(node.GetValue("cond"))
            Set bdy = node.GetValue("body")
            Set nbdy = New Collection
            If Not bdy Is Nothing Then
                Dim bi As Long
                For bi = 1 To bdy.Count
                    nbdy.Add NormalizeNodeRecursive(bdy(bi))
                Next bi
            End If
            Set nw = New Map
            nw.Add "type", "While"
            nw.SetValue "cond", IIf(cnd Is Nothing, node.GetValue("cond"), cnd)
            nw.SetValue "body", nbdy
            Set NormalizeNodeRecursive = nw
            Exit Function

        Case "TryCatch"
            Set tryBlk = node.GetValue("try")
            Set catchBlk = node.GetValue("catch")
            Set nt = New Collection: Set nc = New Collection
            Dim ti As Long
            If Not tryBlk Is Nothing Then
                For ti = 1 To tryBlk.Count: nt.Add NormalizeNodeRecursive(tryBlk(ti)): Next ti
            End If
            If Not catchBlk Is Nothing Then
                For ti = 1 To catchBlk.Count: nc.Add NormalizeNodeRecursive(catchBlk(ti)): Next ti
            End If
            Set ntc = New Map
            ntc.Add "type", "TryCatch"
            ntc.SetValue "try", nt
            ntc.SetValue "catch", nc
            Set NormalizeNodeRecursive = ntc
            Exit Function

        Case "Switch"
            Set cases = node.GetValue("cases")
            Set ncases = New Collection
            If Not cases Is Nothing Then
                Dim ci As Long
                For ci = 1 To cases.Count
                    Set pair = cases(ci)
                    Set caseExpr = NormalizeExprNodeRecursive(pair(1))
                    Set blockStmts = pair(2)
                    Set nb = New Collection
                    Dim bi2 As Long
                    For bi2 = 1 To blockStmts.Count
                        nb.Add NormalizeNodeRecursive(blockStmts(bi2))
                    Next bi2
                    Set newPair = New Collection
                    newPair.Add caseExpr
                    newPair.Add nb
                    ncases.Add newPair
                Next ci
            End If
            Set def = node.GetValue("default")
            If Not def Is Nothing Then
                Set ndef = New Collection
                Dim di As Long
                For di = 1 To def.Count: ndef.Add NormalizeNodeRecursive(def(di)): Next di
            End If
            Set nsw = New Map
            nsw.Add "type", "Switch"
            nsw.SetValue "expr", NormalizeExprNodeRecursive(node.GetValue("expr"))
            nsw.SetValue "cases", ncases
            If Not ndef Is Nothing Then nsw.SetValue "default", ndef
            Set NormalizeNodeRecursive = nsw
            Exit Function

        Case Else
            ' Assign, Print, Return, Break, Continue, etc. often have expr fields to normalize.
            If node.GetValue("type") = "Assign" Then
                Set leftA = node.GetValue("left")
                Set rightA = NormalizeExprNodeRecursive(node.GetValue("right"))
                Set naNode = New Map
                naNode.Add "type", "Assign"
                naNode.SetValue "left", leftA
                naNode.SetValue "right", IIf(rightA Is Nothing, node.GetValue("right"), rightA)
                Set NormalizeNodeRecursive = naNode
                Exit Function
            End If
            Set NormalizeNodeRecursive = node
            Exit Function
    End Select
End Function

Private Function NormalizeAssignsInStmts(stmts As Collection) As Collection
    Dim out As New Collection
    If stmts Is Nothing Then
        Set NormalizeAssignsInStmts = out: Exit Function
    End If
    Dim i As Long
    Dim n As Map
    For i = 1 To stmts.Count
        Set n = stmts(i)
        out.Add NormalizeNodeRecursive(n)
    Next i
    Set NormalizeAssignsInStmts = out
End Function

 ' Normalize compound assignments like a += b  -> Assign(a, Binary(a, +, b))
 Private Function NormalizeCompoundAssigns(stmts As Collection) As Collection
     Dim out As New Collection
     If stmts Is Nothing Then
         Set NormalizeCompoundAssigns = out: Exit Function
     End If
     Dim i As Long
     Dim e As Map
     Dim n As Map
     Dim op As String
     Dim left_ As Map
     Dim right_ As Map
     Dim bin As Map
     Dim baseOp As String
     Dim a As Map
     For i = 1 To stmts.Count
         Set n = stmts(i)
         If Not n Is Nothing And n.GetValue("type") = "ExprStmt" Then
             Set e = n.GetValue("expr")
             If Not e Is Nothing And e.GetValue("type") = "Binary" Then
                 op = CStr(e.GetValue("op"))
                 Select Case op
                     Case "+=", "-=", "*=", "/=", "%=", "^="
                         ' build Assign node: left = left <op> right
                         Set left_ = e.GetValue("left")
                         Set right_ = e.GetValue("right")
                         Set bin = New Map
                         bin.Add "type", "Binary"
                         baseOp = left$(op, Len(op) - 1)
                         bin.SetValue "op", baseOp
                         ' left needs to be used as a copy for the RHS binary left operand
                         bin.SetValue "left", left_
                         bin.SetValue "right", right_
                         Set a = New Map
                         a.Add "type", "Assign"
                         a.SetValue "left", left_
                         a.SetValue "right", bin
                         out.Add a
                         GoTo NormNext
                 End Select
             End If
         End If
         ' default: pass-through (but recursively normalize inner blocks too if desired)
         out.Add n
NormNext:
     Next i
     Set NormalizeCompoundAssigns = out
 End Function
' ---------------------
' End Compiler class
' ---------------------



