Attribute VB_Name = "FormulasTest"
Option Explicit

Sub TestAll()
    TestTokenize
    TestPretty
End Sub

Sub TestTokenize()
    Dim tests As Collection
    Set tests = New Collection
    tests.Add Array( _
        "tokenize math operators", _
        "+-*/", _
        Stringify(Array( _
            Token(TK_PUNCT, "+", 1), Token(TK_PUNCT, "-", 2), Token(TK_PUNCT, "*", 3), Token(TK_PUNCT, "/", 4) _
        )))
    tests.Add Array( _
        "tokenize parentheses", _
        "()", _
        Stringify(Array( _
            Token(TK_PUNCT, "(", 1), Token(TK_PUNCT, ")", 2) _
        )))
    tests.Add Array( _
        "tokenize ident", _
        "var,\a,FU_NC", _
        Stringify(Array( _
            Token(TK_IDENT, "var", 1), _
            Token(TK_PUNCT, ",", 4), _
            Token(TK_IDENT, "\a", 5), _
            Token(TK_PUNCT, ",", 7), _
            Token(TK_IDENT, "FU_NC", 8) _
        )))
    tests.Add Array( _
        "tokenize simple function call", _
        "SUM(12)", _
        Stringify(Array( _
            Token(TK_FUNCNAME, "SUM", 1), Token(TK_PUNCT, "(", 4), Token(TK_NUM, "12", 5), Token(TK_PUNCT, ")", 7) _
        )))
    tests.Add Array( _
        "tokenize multi-arg function call", _
        "SUM(12, 34)", _
        Stringify(Array( _
            Token(TK_FUNCNAME, "SUM", 1), Token(TK_PUNCT, "(", 4), _
            Token(TK_NUM, "12", 5), Token(TK_PUNCT, ",", 7), Token(TK_NUM, "34", 9), _
            Token(TK_PUNCT, ")", 11) _
        )))
    tests.Add Array( _
        "tokenize comparison operators", _
        "1=2<>3<4<=5>6>=7", _
        Stringify(Array( _
            Token(TK_NUM, "1", 1), Token(TK_PUNCT, "=", 2), Token(TK_NUM, "2", 3), _
            Token(TK_PUNCT, "<>", 4), Token(TK_NUM, "3", 6), _
            Token(TK_PUNCT, "<", 7), Token(TK_NUM, "4", 8), _
            Token(TK_PUNCT, "<=", 9), Token(TK_NUM, "5", 11), _
            Token(TK_PUNCT, ">", 12), Token(TK_NUM, "6", 13), _
            Token(TK_PUNCT, ">=", 14), Token(TK_NUM, "7", 16) _
        )))
    tests.Add Array( _
        "tokenize nested function call", _
        "SUM(MIN(1, MAX(3, NOW())))", _
        Stringify(Array( _
            Token(TK_FUNCNAME, "SUM", 1), Token(TK_PUNCT, "(", 4), _
                Token(TK_FUNCNAME, "MIN", 5), Token(TK_PUNCT, "(", 8), Token(TK_NUM, "1", 9), Token(TK_PUNCT, ",", 10), _
                    Token(TK_FUNCNAME, "MAX", 12), Token(TK_PUNCT, "(", 15), Token(TK_NUM, "3", 16), Token(TK_PUNCT, ",", 17), _
                        Token(TK_FUNCNAME, "NOW", 19), Token(TK_PUNCT, "(", 22), Token(TK_PUNCT, ")", 23), _
                    Token(TK_PUNCT, ")", 24), _
                Token(TK_PUNCT, ")", 25), _
            Token(TK_PUNCT, ")", 26) _
        )))
    tests.Add Array( _
        "tokenize string literas", _
        """a b c""", _
        Stringify(Array( _
            Token(TK_STRING, """a b c""", 1) _
        )))
    tests.Add Array( _
        "tokenize concatenation", _
        "(+1&""abc"")&NOW()", _
        Stringify(Array( _
            Token(TK_PUNCT, "(", 1), _
            Token(TK_PUNCT, "+", 2), _
            Token(TK_NUM, "1", 3), _
            Token(TK_PUNCT, "&", 4), _
            Token(TK_STRING, """abc""", 5), _
            Token(TK_PUNCT, ")", 10), _
            Token(TK_PUNCT, "&", 11), _
            Token(TK_FUNCNAME, "NOW", 12), _
            Token(TK_PUNCT, "(", 15), _
            Token(TK_PUNCT, ")", 16) _
    )))
    tests.Add Array( _
        "tokenize array", _
        "{1,2;""3"",""4"";TRUE,FALSE}*{2,2}", _
        Stringify(Array( _
            Token(TK_PUNCT, "{", 1), _
            Token(TK_NUM, "1", 2), _
            Token(TK_PUNCT, ",", 3), _
            Token(TK_NUM, "2", 4), _
            Token(TK_PUNCT, ";", 5), _
            Token(TK_STRING, """3""", 6), _
            Token(TK_PUNCT, ",", 9), _
            Token(TK_STRING, """4""", 10), _
            Token(TK_PUNCT, ";", 13), _
            Token(TK_IDENT, "TRUE", 14), _
            Token(TK_PUNCT, ",", 18), _
            Token(TK_IDENT, "FALSE", 19), _
            Token(TK_PUNCT, "}", 24), _
            Token(TK_PUNCT, "*", 25), _
            Token(TK_PUNCT, "{", 26), _
            Token(TK_NUM, "2", 27), _
            Token(TK_PUNCT, ",", 28), _
            Token(TK_NUM, "2", 29), _
            Token(TK_PUNCT, "}", 30) _
        )))
    tests.Add Array( _
        "tokenize address and ident", _
        "a1 A1 XFD1 XFE1 A1048576 A1048577", _
        Stringify(Array( _
            Token(TK_REF, "a1", 1), _
            Token(TK_REF, "A1", 4), _
            Token(TK_REF, "XFD1", 7), _
            Token(TK_IDENT, "XFE1", 12), _
            Token(TK_REF, "A1048576", 17), _
            Token(TK_IDENT, "A1048577", 26) _
        )))
    Dim t As Variant
    For Each t In tests
        If IsArray(t) Then
            RunTokenizeTest t
        End If
    Next t
    Debug.Print
End Sub

Private Sub RunTokenizeTest(t As Variant)
    On Error GoTo Catch
        Dim actual As String
        actual = JsonConverter.ConvertToJson(Application.Run("Formulas.Tokenize", CStr(t(1))))
        If actual = CStr(t(2)) Then
            Debug.Print "ok: " & t(0)
        Else
            Debug.Print "ng: " & t(0)
            Debug.Print "assert failed: "
            Debug.Print "  " & "input: " & t(1)
            Debug.Print "  " & "left  == " & actual
            Debug.Print "  " & "right == " & t(2)
            Debug.Print
        End If
Catch:
    If Err.Number <> 0 Then
        If Left(t(0), 6) = "failed" Then
            Debug.Print "ok: " & t(0)
        Else
            Debug.Print "ng: " & t(0)
            Debug.Print "  " & Err.Description
        End If
    End If
End Sub

Sub TestParse()
    Dim tests As Variant
    tests = Array( _
        "=1+2", _
        "=1+2*3", _
        "=(1+2)*3", _
        "=x+y*z", _
        "=(ab+cd)*ef", _
        "=+12*-3/+xyz", _
        "=1=2<>3<4<=5>6>=7", _
        "=(((((1=2)<>3)<4)<=5)>6)>=7", _
        "=SUM(1,2)", _
        "=SUM.1(MIN(a))", _
        "=IF(AND(1=1,MIN(x)=MAX(y)),NOW(),DATE(1990,1,1))", _
        "=""a b c""", _
        "=(+1&""abc"")&NOW()", _
        "={1,2;""3"",""4"";TRUE,FALSE}*{2,2}", _
        "=SUM(A1, B1:C1, (D1:E1:F1))", _
        "" _
    )
    Dim fmt As Formulas.Formatter
    fmt = Formulas.NewFormatter( _
        indent:=" ", _
        indentLength:=2, _
        newLine:=vbCrLf, _
        eqAtStart:=False, _
        newLineAtEof:=False _
    )
    Dim t As Variant
    For Each t In tests
        If CStr(t) <> "" Then
            Debug.Print t
            Debug.Print Formulas.DebugAst(Formulas.Parse(CStr(t)), fmt)
            Debug.Print
        End If
    Next t
End Sub

Sub TestPretty()
    Dim tests As Collection
    Set tests = New Collection
    tests.Add Array( _
        "pretty simple addition", _
        "=1+2", _
        "1 + 2" _
    )
    tests.Add Array( _
        "pretty parentheses", _
        "=(1+2)*3", _
        "(" & vbCrLf & _
        "  1 + 2" & vbCrLf & _
        ") * 3" _
    )
    tests.Add Array( _
        "pretty function with args", _
        "=SUM(A1, B1:C1, (D1:E1:F1))", _
        "SUM(" & vbCrLf & _
        "  A1," & vbCrLf & _
        "  B1:C1," & vbCrLf & _
        "  (" & vbCrLf & _
        "    D1:E1:F1" & vbCrLf & _
        "  )" & vbCrLf & _
        ")" _
    )
    tests.Add Array( _
        "pretty nested function", _
        "=SUM(MIN(1, MAX(3, NOW())))", _
        "SUM(" & vbCrLf & _
        "  MIN(" & vbCrLf & _
        "    1," & vbCrLf & _
        "    MAX(" & vbCrLf & _
        "      3," & vbCrLf & _
        "      NOW()" & vbCrLf & _
        "    )" & vbCrLf & _
        "  )" & vbCrLf & _
        ")" _
    )
    tests.Add Array( _
        "pretty string literal", _
        "=(+1&""abc"")&NOW()", _
        "(" & vbCrLf & _
        "  1 & ""abc""" & vbCrLf & _
        ") & NOW()" _
    )
    tests.Add Array( _
        "pretty concatenation", _
        "=""hello world""", _
        """hello world""" _
    )
    tests.Add Array( _
        "pretty function comparison", _
        "=MIN(x)=MAX(y)", _
        "MIN(" & vbCrLf & _
        "  x" & vbCrLf & _
        ") = MAX(" & vbCrLf & _
        "  y" & vbCrLf & _
        ")" _
    )
    tests.Add Array( _
        "pretty complex expression", _
        "=IF(AND(1=1,MIN(x)=MAX(y)),NOW(),DATE(1990,1,1))", _
        "IF(" & vbCrLf & _
        "  AND(" & vbCrLf & _
        "    1 = 1," & vbCrLf & _
        "    MIN(" & vbCrLf & _
        "      x" & vbCrLf & _
        "    ) = MAX(" & vbCrLf & _
        "      y" & vbCrLf & _
        "    )" & vbCrLf & _
        "  )," & vbCrLf & _
        "  NOW()," & vbCrLf & _
        "  DATE(" & vbCrLf & _
        "    1990," & vbCrLf & _
        "    1," & vbCrLf & _
        "    1" & vbCrLf & _
        "  )" & vbCrLf & _
        ")" _
    )
    tests.Add Array( _
        "pretty simple array", _
        "={1,2}", _
        "{" & vbCrLf & _
        "  1, 2" & vbCrLf & _
        "}" _
    )
    tests.Add Array( _
        "pretty array in func", _
        "=LET(arr,{1,2;3,4},arr)", _
        "LET(" & vbCrLf & _
        "  arr," & vbCrLf & _
        "  {" & vbCrLf & _
        "    1, 2;" & vbCrLf & _
        "    3, 4" & vbCrLf & _
        "  }," & vbCrLf & _
        "  arr" & vbCrLf & _
        ")" _
    )
    tests.Add Array( _
        "pretty function with omitted args", _
        "=SUM(1,,2,,)", _
        "SUM(" & vbCrLf & _
        "  1," & vbCrLf & _
        "  ," & vbCrLf & _
        "  2," & vbCrLf & _
        "  ," & vbCrLf & _
        "  " & vbCrLf & _
        ")" _
    )
    tests.Add Array( _
        "failed parse", _
        "={a}", _
        "}" _
    )
    tests.Add Array( _
        "pretty array", _
        "={1,2;""3"",""4"";TRUE,FALSE}*{2,2}", _
        "{" & vbCrLf & _
        "  1, 2;" & vbCrLf & _
        "  ""3"", ""4"";" & vbCrLf & _
        "  TRUE, FALSE" & vbCrLf & _
        "} * {" & vbCrLf & _
        "  2, 2" & vbCrLf & _
        "}" _
    )
    tests.Add Array( _
        "pretty functions", _
        "=CONCAT(""R"",MOD(ROW()-6,2)*2+1,""C"",INT((ROW()-6)/2)*2+1)", _
        "CONCAT(" & vbCrLf & _
        "  ""R""," & vbCrLf & _
        "  MOD(" & vbCrLf & _
        "    ROW() - 6," & vbCrLf & _
        "    2" & vbCrLf & _
        "  ) * 2 + 1," & vbCrLf & _
        "  ""C""," & vbCrLf & _
        "  INT(" & vbCrLf & _
        "    (" & vbCrLf & _
        "      ROW() - 6" & vbCrLf & _
        "    ) / 2" & vbCrLf & _
        "  ) * 2 + 1" & vbCrLf & _
        ")" _
    )
    Dim fmt As Formulas.Formatter
    fmt = Formulas.NewFormatter( _
        indent:=" ", _
        indentLength:=2, _
        newLine:=vbCrLf, _
        eqAtStart:=False, _
        newLineAtEof:=False _
    )
    Dim t As Variant
    For Each t In tests
        If IsArray(t) Then
            RunPrettyTest t, fmt
        End If
    Next t
    Debug.Print
End Sub

Private Sub RunPrettyTest(t As Variant, fmt As Formulas.Formatter)
    On Error GoTo Catch
        Dim actual As String
        actual = Formulas.Pretty(CStr(t(1)), fmt)
        If actual = CStr(t(2)) Then
            Debug.Print "ok: " & t(0)
        Else
            Debug.Print "ng: " & t(0)
            Debug.Print "assert failed: "
            Debug.Print "  " & "input: " & t(1)
            Debug.Print "  " & "left  == " & actual
            Debug.Print "  " & "right == " & t(2)
            Debug.Print
        End If
Catch:
    If Err.Number <> 0 Then
        If Left(t(0), 6) = "failed" Then
            Debug.Print "ok: " & t(0)
        Else
            Debug.Print "ng: " & t(0)
            Debug.Print "  " & Err.Description
        End If
    End If
End Sub

Private Function Token(kind As TokenKind, val As String, col As Long) As Variant()
    Token = Array(kind, val, col)
End Function

Private Function Stringify(val As Variant) As String
    Stringify = JsonConverter.ConvertToJson(val)
End Function

Private Sub DumpNode(node As Dictionary, indentLevel As Long)
    Dim k As NodeKind
    k = node("kind")
    Dim indent As String
    Dim prefix As String
    indent = String(indentLevel * 2, " ")
    prefix = indentLevel & " " & indent
    If node.Exists("enclosed") Then
        Debug.Print prefix & "- " & "enclosed: " & node("enclosed")
    End If
    Select Case k
        Case Formulas.NodeKind.ND_NUM, Formulas.NodeKind.ND_IDENT
            Debug.Print prefix & "- " & "kind: " & k
            Debug.Print prefix & "- " & "val: " & node("val")
        Case Formulas.NodeKind.ND_ADD, Formulas.NodeKind.ND_SUB, Formulas.NodeKind.ND_MUL, Formulas.NodeKind.ND_DIV, _
             Formulas.NodeKind.ND_EQ, Formulas.NodeKind.ND_NE, _
             Formulas.NodeKind.ND_LT, Formulas.NodeKind.ND_LE, Formulas.NodeKind.ND_GT, Formulas.NodeKind.ND_GE, _
             Formulas.NodeKind.ND_CONCAT
            Debug.Print prefix & "- " & "kind: " & k
            Debug.Print prefix & "lhs:"
            Call DumpNode(node("lhs"), indentLevel + 1)
            Debug.Print prefix & "rhs:"
            Call DumpNode(node("rhs"), indentLevel + 1)
        Case Formulas.NodeKind.ND_FUNC
            Debug.Print prefix & "- " & "kind: " & k
            Debug.Print prefix & "- " & "name: " & node("name")
            Debug.Print prefix & "- " & "args:"
            Dim x As Dictionary
            For Each x In node("args")
                Call DumpNode(x, indentLevel + 1)
            Next x
        Case Formulas.NodeKind.ND_STRING
            Debug.Print prefix & "- " & "kind: " & k
            Debug.Print prefix & "- " & "val: " & Chr(34) & node("val") & Chr(34)
        Case Formulas.NodeKind.ND_ARRAY
            Debug.Print prefix & "- " & "kind: " & k
            Debug.Print prefix & "- " & "elements:"
            Dim r As Dictionary
            For Each r In node("elements")
                DumpNode r, indentLevel + 1
            Next r
        Case Formulas.NodeKind.ND_ARRAY_ROW
            Debug.Print prefix & "- " & "kind: " & k
            Debug.Print prefix & "- " & "elements:"
            Dim c As Dictionary
            For Each c In node("elements")
                DumpNode c, indentLevel
            Next c
    End Select
    If indentLevel = 0 Then
        Debug.Print
    End If
End Sub
