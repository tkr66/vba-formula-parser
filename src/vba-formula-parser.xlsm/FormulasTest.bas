Attribute VB_Name = "FormulasTest"
Option Explicit

Sub TestTokenize()
    Dim tests As Collection
    Set tests = New Collection
    tests.Add Array( _
        "test tokenize math operators", _
        "+-*/", _
        Stringify(Array( _
            Token(TK_PUNCT, "+", 1), Token(TK_PUNCT, "-", 2), Token(TK_PUNCT, "*", 3), Token(TK_PUNCT, "/", 4) _
        )))
    tests.Add Array( _
        "test tokenize parentheses", _
        "()", _
        Stringify(Array( _
            Token(TK_PUNCT, "(", 1), Token(TK_PUNCT, ")", 2) _
        )))
    tests.Add Array( _
        "test tokenize simple function call", _
        "SUM(12)", _
        Stringify(Array( _
            Token(TK_FUNCNAME, "SUM", 1), Token(TK_PUNCT, "(", 4), Token(TK_NUM, "12", 5), Token(TK_PUNCT, ")", 7) _
        )))
    tests.Add Array( _
        "test tokenize multi-arg function call", _
        "SUM(12, 34)", _
        Stringify(Array( _
            Token(TK_FUNCNAME, "SUM", 1), Token(TK_PUNCT, "(", 4), _
            Token(TK_NUM, "12", 5), Token(TK_PUNCT, ",", 7), Token(TK_NUM, "34", 9), _
            Token(TK_PUNCT, ")", 11) _
        )))
    tests.Add Array( _
        "test tokenize comparison operators", _
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
        "test tokenize nested function call", _
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
        "test tokenize string literas", _
        """a b c""", _
        Stringify(Array( _
            Token(TK_STRING, "a b c", 1) _
        )))
    Dim t As Variant
    For Each t In tests
        If IsArray(t) Then
            Dim json As String
            json = JsonConverter.ConvertToJson(Formulas.Tokenize(CStr(t(1))))
            If CStr(json) = CStr(t(2)) Then
                Debug.Print "ok: " & t(0)
            Else
                Debug.Print "assert failed: " & t(0)
                Debug.Print "  " & "input: " & t(1)
                Debug.Print "  " & "left  == " & json
                Debug.Print "  " & "right == " & t(2)
                Debug.Print
            End If
        End If
    Next t
    Debug.Print
End Sub

Sub TestParse()
    Dim tests As Variant
    tests = Array( _
        "1+2", _
        "1+2*3", _
        "(1+2)*3", _
        "x+y*z", _
        "(ab+cd)*ef", _
        "+12*-3/+xyz", _
        "1=2<>3<4<=5>6>=7", _
        "(((((1=2)<>3)<4)<=5)>6)>=7", _
        "SUM(1,2)", _
        "SUM.1(MIN(a))", _
        "IF(AND(1=1,MIN(x)=MAX(y)),NOW(),DATE(1990,1,1))", _
        """a b c""", _
        "" _
    )
    Dim t As Variant
    For Each t In tests
        If CStr(t) <> "" Then
            Debug.Print t
            Call DumpNode(Formulas.Parse(CStr(t)), 0)
        End If
    Next t
End Sub

Sub TestPretty()
    Dim tests As Collection
    Set tests = New Collection
    tests.Add Array( _
        "test pretty simple addition", _
        "1+2", _
        "1 + 2" _
    )
    tests.Add Array( _
        "test pretty parentheses", _
        "(1+2)*3", _
        "(" & vbCrLf & _
        "  1 + 2" & vbCrLf & _
        ") * 3" _
    )
    tests.Add Array( _
        "test pretty function with args", _
        "SUM(A, B)", _
        "SUM(" & vbCrLf & _
        "  A," & vbCrLf & _
        "  B" & vbCrLf & _
        ")" _
    )
    tests.Add Array( _
        "test pretty nested function", _
        "SUM(MIN(1, MAX(3, NOW())))", _
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
        "test pretty string literal", _
        """hello world""", _
        """hello world""" _
    )
    tests.Add Array( _
        "test pretty complex expression", _
        "IF(AND(1=1,MIN(x)=MAX(y)),NOW(),DATE(1990,1,1))", _
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
    Dim t As Variant
    For Each t In tests
        If IsArray(t) Then
            Dim actualPretty As String
            actualPretty = Formulas.Pretty(Formulas.Parse(CStr(t(1))), 2)
            If actualPretty = CStr(t(2)) Then
                Debug.Print "ok: " & t(0)
            Else
                Debug.Print "assert failed: " & t(0)
                Debug.Print "  " & "input: " & t(1)
                Debug.Print "  " & "left  == " & actualPretty
                Debug.Print "  " & "right == " & t(2)
                Debug.Print
            End If
        End If
    Next t
    Debug.Print
End Sub

Private Function Token(kind As TokenKind, val As String, col As Long) As Variant()
    Token = Array(kind, val, col)
End Function

Private Function Stringify(val As Variant) As String
    Stringify = JsonConverter.ConvertToJson(val)
End Function

Private Function TokenKindMap() As Dictionary
    Set TokenKindMap = Formulas.TokenKindMap
End Function

Private Function NodeKindMap() As Dictionary
    Set NodeKindMap = Formulas.NodeKindMap
End Function

Private Sub DumpTokens(toks As Collection)
    Dim t As Variant
    For Each t In toks
        Debug.Print "kind: " & TokenKindMap(t(0)) & ", val: " & t(1)
    Next t
    Debug.Print
End Sub

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
            Debug.Print prefix & "- " & "kind: " & NodeKindMap(k)
            Debug.Print prefix & "- " & "val: " & node("val")
        Case Formulas.NodeKind.ND_ADD, Formulas.NodeKind.ND_SUB, Formulas.NodeKind.ND_MUL, Formulas.NodeKind.ND_DIV, _
             Formulas.NodeKind.ND_EQ, Formulas.NodeKind.ND_NE, _
             Formulas.NodeKind.ND_LT, Formulas.NodeKind.ND_LE, Formulas.NodeKind.ND_GT, Formulas.NodeKind.ND_GE
            Debug.Print prefix & "- " & "kind: " & NodeKindMap(k)
            Debug.Print prefix & "lhs:"
            Call DumpNode(node("lhs"), indentLevel + 1)
            Debug.Print prefix & "rhs:"
            Call DumpNode(node("rhs"), indentLevel + 1)
        Case Formulas.NodeKind.ND_FUNC
            Debug.Print prefix & "- " & "kind: " & NodeKindMap(k)
            Debug.Print prefix & "- " & "name: " & node("name")
            Debug.Print prefix & "- " & "args:"
            Dim x As Dictionary
            For Each x In node("args")
                Call DumpNode(x, indentLevel + 1)
            Next x
        Case Formulas.NodeKind.ND_STRING
            Debug.Print prefix & "- " & "kind: " & NodeKindMap(k)
            Debug.Print prefix & "- " & "val: " & Chr(34) & node("val") & Chr(34)
    End Select
    If indentLevel = 0 Then
        Debug.Print
    End If
End Sub
