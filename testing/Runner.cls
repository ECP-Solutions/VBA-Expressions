VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Runner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Module: Runner
Option Explicit

Public Sub RunExample()
    ASF_InitGlobals
    Dim src As String
    src = _
      "/* demo */" & vbCrLf & _
      "print('hello ASF')," & vbCrLf & _
      "a = 2, b = 3, print(a + b), " & vbCrLf & _
      "if (a < b) { print('a is less') } else { print('a >= b') }," & vbCrLf & _
      "s = 'it\'s OK', print(s)," & vbCrLf & _
      "for(i=1,i<=3,i=i+1) { print(i) }, " & vbCrLf & _
      "x = true, y = false, print(x && y), print(x || y), print(!y)," & vbCrLf & _
      "switch(a) { case 1 { print('one') } case 2 { print('two') break } default { print('other') } }"
    Dim c As New Compiler
    Dim idx As Long: idx = c.CompileProgram(src, "example_prog")
    Dim VM As New VM
    VM.RunProgramByIndex idx
    Debug.Print "=== Runtime Log ==="
    Dim i As Long
    For i = 1 To gRuntimeLog.Count
        Debug.Print gRuntimeLog(i)
    Next i
End Sub


Public Sub RunExample2()
    ASF_InitGlobals
    Dim src As String
'    src = "a=2; if (a==1) { print('one') } elseif (a==2) { print('two') } else { print('other') }; print('done')" 'PASSED
'    src = "x=false; print(x && (1/0))" 'PASSED
'    src = "x=true; print(x || (1/0))" 'PASSED
'    src = "a=3;" & vbCrLf & _
'                   "if (a=1) {" & vbCrLf & "  print('one')" & vbCrLf & _
'                   "} elseif (a=2) {" & vbCrLf & "  print('two')" & vbCrLf & _
'                   "} elseif (a=3) {" & vbCrLf & "  print('three')" & vbCrLf & _
'                   "} else {" & vbCrLf & "  print('other')" & vbCrLf & _
'                   "}; print('end')"  'PASSED
'    src = "s=0; for(i=1,i<=5,i=i+1) { if (i=3) { continue } if (i=5) { break } s = s + i }; print(s)" 'PASSED
'    src = "i=1;s=0; while (i<=5) { if (i=2) { i = i + 1 : continue } if (i=5) { break }; s = s + i ; i = i + 1 }; print(s)" 'PASSED
    src = "a = 1; f = fun() { a = a + 1; return a; }; print(f()); print(a)"
    Dim c As New Compiler
    Dim idx As Long: idx = c.CompileProgram(src, "example_prog")
    Dim VM As New VM
    VM.RunProgramByIndex idx
    Debug.Print "=== Runtime Log ==="
    Dim i As Long
    For i = 1 To gRuntimeLog.Count
        Debug.Print gRuntimeLog(i)
    Next i
End Sub

