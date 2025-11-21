VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Parser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Module: Parser
Option Explicit
Private Const d_Apostrophe As String = "'"

Private GLOBALS_ As Globals

Public Sub SetGlobals(ByRef aGlobals As Globals)
    Set GLOBALS_ = aGlobals
End Sub

Public Function Tokenize(src As String) As Collection
    Dim toks As New Collection
    Dim i As Long, n As Long
    Dim ch As String
    Dim nxt As String
    Dim j As Long, jC As Long, jB As Long
    n = Len(src)
    i = 1

Tokenize_MainLoop:
    Do While i <= n
        If i > n Then Exit Do ' FIX: Bounds
        ch = Mid$(src, i, 1)
        ' comments and slash handling
        If ch = "/" And i < n Then
            nxt = Mid$(src, i + 1, 1)
            If nxt = "/" Then
                jC = i + 2
                Do While jC <= n And Mid$(src, jC, 1) <> vbCr And Mid$(src, jC, 1) <> vbLf
                    jC = jC + 1
                Loop
                toks.Add Array("COMMENT", Mid$(src, i + 2, jC - i - 2))
                i = jC
                GoTo Tokenize_SkipToLoopEnd
            ElseIf nxt = "*" Then
                jB = i + 2
                Dim foundClose As Boolean: foundClose = False
                Do While jB <= n - 1
                    If Mid$(src, jB, 2) = "*/" Then
                        foundClose = True
                        Exit Do
                    End If
                    jB = jB + 1
                Loop
                If foundClose Then
                    toks.Add Array("COMMENT", Mid$(src, i + 2, jB - i - 2))
                    i = jB + 2
                    GoTo Tokenize_SkipToLoopEnd
                Else
                    toks.Add Array("COMMENT", Mid$(src, i + 2))
                    i = n + 1
                    GoTo Tokenize_SkipToLoopEnd
                End If
            End If
        End If

        If ch = "#" Then
            Dim jH As Long: jH = i + 1
            Do While jH <= n And Mid$(src, jH, 1) <> vbCr And Mid$(src, jH, 1) <> vbLf
                jH = jH + 1
            Loop
            toks.Add Array("COMMENT", Mid$(src, i + 1, jH - i - 1))
            i = jH
            GoTo Tokenize_SkipToLoopEnd
        End If

        Select Case ch
            Case " ", vbTab, vbCr, vbLf
                i = i + 1
                GoTo Tokenize_SkipToLoopEnd
            Case ","
                toks.Add Array("SEP", ch): i = i + 1: GoTo Tokenize_SkipToLoopEnd
            Case ";"
                toks.Add Array("SEP", ch): i = i + 1: GoTo Tokenize_SkipToLoopEnd
            Case "(", ")", "{", "}", "[", "]"
                toks.Add Array("PAREN", ch): i = i + 1: GoTo Tokenize_SkipToLoopEnd
            Case "+", "-", "*", "/", "^", "%", "&", "|"
                ' support +=, -=, *=, /=, ^=, %=, &=, |=
                If i < n And Mid$(src, i + 1, 1) = "=" Then
                    toks.Add Array("OP", ch & "="): i = i + 2
                Else
                    toks.Add Array("OP", ch): i = i + 1
                End If
                GoTo Tokenize_SkipToLoopEnd
            Case "&"
                If i < n And Mid$(src, i + 1, 1) = "&" Then toks.Add Array("OP", "&&"): i = i + 2 Else toks.Add Array("OP", "&"): i = i + 1
                GoTo Tokenize_SkipToLoopEnd
            Case "|"
                If i < n And Mid$(src, i + 1, 1) = "|" Then toks.Add Array("OP", "||"): i = i + 2 Else toks.Add Array("SYM", "|"): i = i + 1
                GoTo Tokenize_SkipToLoopEnd
            Case "!"
                If i < n And Mid$(src, i + 1, 1) = "=" Then toks.Add Array("OP", "!="): i = i + 2 Else toks.Add Array("OP", "!"): i = i + 1
                GoTo Tokenize_SkipToLoopEnd
            Case "<", ">"
                Dim two As String
                Dim three As String
                If i < n Then two = Mid$(src, i, 2) Else two = ""
                If two = ">=" Or two = "<=" Then
                    toks.Add Array("OP", two): i = i + 2
                Else
                    If i < n Then three = Mid$(src, i, 3) Else three = ""
                    If three = "<<=" Or three = ">>=" Then 'Support for compound shift assignment
                        toks.Add Array("OP", three): i = i + 3
                    Else
                        If two = "<<" Or two = ">>" Then 'Support for eleft and right shifts
                            toks.Add Array("OP", two): i = i + 2
                        Else
                            toks.Add Array("OP", ch): i = i + 1
                        End If
                    End If
                End If
                GoTo Tokenize_SkipToLoopEnd
            Case "="
                If i < n And Mid$(src, i + 1, 1) = "=" Then toks.Add Array("OP", "=="): i = i + 2 Else toks.Add Array("OP", "="): i = i + 1
                GoTo Tokenize_SkipToLoopEnd
            Case "?"
                toks.Add Array("OP", ch): i = i + 1: GoTo Tokenize_SkipToLoopEnd
            Case ":"
                toks.Add Array("OP", ch): i = i + 1: GoTo Tokenize_SkipToLoopEnd
            Case "'"
                j = i + 1
                Dim sb As String: sb = ""
                Do While j <= n
                    Dim cj As String: cj = Mid$(src, j, 1)
                    If cj = "\" Then
                        If j + 1 <= n Then
                            Dim nx As String: nx = Mid$(src, j + 1, 1)
                            Select Case nx
                                Case "'": sb = sb & "'": j = j + 2: GoTo ContinueDo
                                Case "\": sb = sb & "\": j = j + 2: GoTo ContinueDo
                                Case "n": sb = sb & vbLf: j = j + 2: GoTo ContinueDo
                                Case "r": sb = sb & vbCr: j = j + 2: GoTo ContinueDo
                                Case "t": sb = sb & vbTab: j = j + 2: GoTo ContinueDo
                                Case Else: sb = sb & nx: j = j + 2: GoTo ContinueDo
                            End Select
                        Else
                            sb = sb & "\"
                            j = j + 1
                            GoTo ContinueDo
                        End If
                    ElseIf cj = "'" Then
                        Exit Do
                    Else
                        sb = sb & cj
                        j = j + 1
                    End If
ContinueDo:
                Loop
                If j <= n And Mid$(src, j, 1) = "'" Then
                    toks.Add Array("STRING", sb)
                    i = j + 1
                    GoTo Tokenize_SkipToLoopEnd
                Else
                    If GLOBALS_.STRICT_STRINGS Then err.Raise vbObjectError + 1001, "Parser.Tokenize", "Unterminated string literal"
                    toks.Add Array("STRING", sb)
                    i = n + 1
                    GoTo Tokenize_SkipToLoopEnd
                End If
            Case Else
                If ch Like "[0-9]" Then
                    Dim j2 As Long: j2 = i
                    Do While j2 <= n And Mid$(src, j2, 1) Like "[0-9.]"
                        j2 = j2 + 1
                    Loop
                    toks.Add Array("NUMBER", Mid$(src, i, j2 - i))
                    i = j2
                    GoTo Tokenize_SkipToLoopEnd
                ElseIf ch Like "[A-Za-z_]" Then
                    Dim j3 As Long: j3 = i
                    Do While j3 <= n And Mid$(src, j3, 1) Like "[A-Za-z0-9_.]"
                        j3 = j3 + 1
                    Loop
                    toks.Add Array("IDENT", Mid$(src, i, j3 - i))
                    i = j3
                    GoTo Tokenize_SkipToLoopEnd
                ' forced VBAexpression: @(...). Capture whole balanced parentheses block as a single VBEXPR token.
                ElseIf ch = "@" And i < n And Mid$(src, i + 1, 1) = "(" Then
                    Dim jvb As Long: jvb = i + 2
                    Dim depthVB As Long: depthVB = 1
                    Dim vbContent As String: vbContent = ""
                    Do While jvb <= n And depthVB > 0
                        Dim cjvb As String: cjvb = Mid$(src, jvb, 1)
                        If cjvb = "(" Then
                            depthVB = depthVB + 1
                            vbContent = vbContent & cjvb
                            jvb = jvb + 1
                        ElseIf cjvb = ")" Then
                            depthVB = depthVB - 1
                            If depthVB = 0 Then
                                ' matched closing ')', stop without adding this ')'
                                Exit Do
                            Else
                                vbContent = vbContent & cjvb
                                jvb = jvb + 1
                            End If
                        Else
                            vbContent = vbContent & cjvb
                            jvb = jvb + 1
                        End If
                    Loop
                    ' vbContent now contains inner characters between parentheses
                    toks.Add Array("VBEXPR", vbContent)
                    i = jvb + 1  ' position after ')'
                    GoTo Tokenize_SkipToLoopEnd
                Else
                    toks.Add Array("SYM", ch)
                    i = i + 1
                    GoTo Tokenize_SkipToLoopEnd
                End If
        End Select

Tokenize_SkipToLoopEnd:
    Loop

    ' collapse array index tokens: IDENT [ NUMBER ] -> IDENT "name[index]"
    Dim out As New Collection
    Dim idx As Long
    idx = 1
Tokenize_CollapseLoop:
    Do While idx <= toks.count
        If idx > toks.count Then Exit Do ' FIX: Bounds
        Dim tk As Variant: tk = toks(idx)
        If tk(0) = "IDENT" And idx + 3 <= toks.count Then
            If toks(idx + 1)(0) = "PAREN" And toks(idx + 1)(1) = "[" And toks(idx + 2)(0) = "NUMBER" And toks(idx + 3)(0) = "PAREN" And toks(idx + 3)(1) = "]" Then
                out.Add Array("IDENT", tk(1) & "[" & toks(idx + 2)(1) & "]")
                idx = idx + 4
                GoTo Tokenize_CollapseLoop
            End If
        End If
        out.Add tk
        idx = idx + 1
    Loop

    Set Tokenize = out
End Function

Public Function JoinTokens(tokens As Collection) As String
    Dim s As String: s = ""
    Dim i As Long
    For i = 1 To tokens.count
        If tokens(i)(0) = "STRING" Then
            s = s & ToLiteralString(CStr(tokens(i)(1))) & " "
        Else
            s = s & tokens(i)(1) & " "
        End If
    Next i
    JoinTokens = Trim$(s)
End Function
Private Function ToLiteralString(ByRef aString As String) As String
    ToLiteralString = d_Apostrophe & aString & d_Apostrophe
End Function
