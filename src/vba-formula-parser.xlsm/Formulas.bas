Attribute VB_Name = "Formulas"
Option Explicit

Private input_ As String
Public Enum TokenKind
    TK_NUM
    TK_PUNCT
    TK_IDENT
    TK_FUNCNAME
End Enum

Private pos_ As Long
Public Enum NodeKind
    ND_NUM
    ND_ADD
    ND_SUB
    ND_MUL
    ND_DIV
    ND_IDENT
    ND_EQ
    ND_NE
    ND_LT
    ND_LE
    ND_GT
    ND_GE
    ND_FUNC
End Enum

Private Const BUF_MAX As Long = 8096

Static Property Get TokenKindMap() As Dictionary
    Set TokenKindMap = New Dictionary
    TokenKindMap.Add TK_NUM, "TK_NUM"
    TokenKindMap.Add TK_PUNCT, "TK_PUNCT"
    TokenKindMap.Add TK_IDENT, "TK_IDENT"
    TokenKindMap.Add TK_FUNCNAME, "TK_FUNCNAME"
End Property

Static Property Get NodeKindMap() As Dictionary
    Set NodeKindMap = New Dictionary
    NodeKindMap.Add ND_NUM, "ND_NUM"
    NodeKindMap.Add ND_ADD, "ND_ADD"
    NodeKindMap.Add ND_SUB, "ND_SUB"
    NodeKindMap.Add ND_MUL, "ND_MUL"
    NodeKindMap.Add ND_DIV, "ND_DIV"
    NodeKindMap.Add ND_IDENT, "ND_IDENT"
    NodeKindMap.Add ND_EQ, "ND_EQ"
    NodeKindMap.Add ND_NE, "ND_NE"
    NodeKindMap.Add ND_LT, "ND_LT"
    NodeKindMap.Add ND_LE, "ND_LE"
    NodeKindMap.Add ND_GT, "ND_GT"
    NodeKindMap.Add ND_GE, "ND_GE"
    NodeKindMap.Add ND_FUNC, "ND_FUNC"
End Property

Static Property Get OperatorMap() As Dictionary
    Set OperatorMap = New Dictionary
    OperatorMap.Add ND_ADD, "+"
    OperatorMap.Add ND_SUB, "-"
    OperatorMap.Add ND_MUL, "*"
    OperatorMap.Add ND_DIV, "/"
    OperatorMap.Add ND_EQ, "="
    OperatorMap.Add ND_NE, "<>"
    OperatorMap.Add ND_LT, "<"
    OperatorMap.Add ND_LE, "<="
    OperatorMap.Add ND_GT, ">"
    OperatorMap.Add ND_GE, ">="
End Property

Public Function Tokenize(str As String) As Collection
    input_ = str
    Dim toks As Collection
    Set toks = New Collection
    Dim start As Long
    Dim i As Long
    i = 1
    Do While i <= Len(str)
        Dim c As String
        c = Mid(str, i, 1)
        Select Case True
            Case c = " "
                i = i + 1
            Case IsNumeric(c)
                start = i
                Do
                    i = i + 1
                Loop While IsNumeric(Mid(str, i, 1))
                toks.Add NewToken(TK_NUM, Mid(str, start, i - start), start)
            Case c = "+" Or c = "-" Or c = "*" Or c = "/"
                toks.Add NewToken(TK_PUNCT, c, i)
                i = i + 1
            Case c = "(" Or c = ")"
                toks.Add NewToken(TK_PUNCT, c, i)
                i = i + 1
            Case c = ","
                toks.Add NewToken(TK_PUNCT, c, i)
                i = i + 1
            Case c = "."
                toks.Add NewToken(TK_PUNCT, c, i)
                i = i + 1
            Case IsIdent(c)
                Dim expectFuncName As Boolean
                Dim cur As String
                start = i
                Do
                    i = i + 1
                    cur = Mid(str, i, 1)
                    Select Case True
                        Case IsIdent(cur)
                        Case cur = "."
                            expectFuncName = True
                        Case IsNumeric(cur)
                        Case Else
                            Exit Do
                    End Select
                Loop
                If expectFuncName Then
                    If Mid(str, i, 1) <> "(" Then
                        Call ErrorAt(str, "expected '('")
                    End If
                    toks.Add NewToken(TK_FUNCNAME, Mid(str, start, i - start), start)
                Else
                    If IsNumeric(Mid(str, i - 1, 1)) Then
                        Call ErrorAt(str, "expected a char")
                    End If
                    If Mid(str, i, 1) = "(" Then
                        toks.Add NewToken(TK_FUNCNAME, Mid(str, start, i - start), start)
                    Else
                        toks.Add NewToken(TK_IDENT, Mid(str, start, i - start), start)
                    End If
                End If
                expectFuncName = False
            Case c = "="
                toks.Add NewToken(TK_PUNCT, c, i)
                i = i + 1
            Case c = "<"
                If Mid(str, i + 1, 1) = ">" Or Mid(str, i + 1, 1) = "=" Then
                    toks.Add NewToken(TK_PUNCT, Mid(str, i, 2), i)
                    i = i + 2
                Else
                    toks.Add NewToken(TK_PUNCT, c, i)
                    i = i + 1
                End If
            Case c = ">"
                If Mid(str, i + 1, 1) = "=" Then
                    toks.Add NewToken(TK_PUNCT, Mid(str, i, 2), i)
                    i = i + 2
                Else
                    toks.Add NewToken(TK_PUNCT, c, i)
                    i = i + 1
                End If
            Case Else
                Call ErrorAt(Mid(str, i), "unexpected token")
        End Select
    Loop
    Set Tokenize = toks
End Function

Private Function NewToken(kind As Long, val As String, col As Long) As Variant()
    NewToken = Array(kind, val, col)
End Function

Private Function IsIdent(c As String) As Boolean
    If c = "" Then
        IsIdent = False
        Exit Function
    End If
    Dim dec As Long
    dec = Asc(c)
    IsIdent = (97 <= dec And dec <= 122) Or (65 <= dec And dec <= 90)
End Function

Private Sub ErrorAt(rest As String, msg As String)
    Dim pos As Long
    pos = Len(input_) - Len(rest)
    Dim prefix As String
    prefix = String(4, " ")
    Debug.Print "tokenize error:"
    Debug.Print prefix & input_
    Debug.Print prefix & String(pos, " ") & "^ " & msg
    Debug.Print
    End
End Sub

Public Function Parse(str As String) As Dictionary
    Dim toks As Collection
    Set toks = Tokenize(str)
    pos_ = 1

    Set Parse = Expr(toks)

    If toks.Count >= pos_ Then
        Call ErrorAt2(toks, "unexpected trailing token")
    End If
End Function

Private Function NewNode(kind As String) As Dictionary
    Set NewNode = New Dictionary
    NewNode.Add "kind", kind
End Function

Private Function NewBinary(kind As String, lhs As Dictionary, rhs As Dictionary) As Dictionary
    Set NewBinary = NewNode(kind)
    NewBinary.Add "lhs", lhs
    NewBinary.Add "rhs", rhs
End Function

Private Function NewNum(val As Long) As Dictionary
    Set NewNum = NewNode(ND_NUM)
    NewNum.Add "val", val
End Function

Private Function NewIdent(val As String) As Dictionary
    Set NewIdent = NewNode(ND_IDENT)
    NewIdent.Add "val", val
End Function

Private Function NewFunc(name_ As String, args_ As Collection) As Dictionary
    Set NewFunc = NewNode(ND_FUNC)
    NewFunc.Add "name", name_
    NewFunc.Add "args", args_
End Function

Private Function Consume(toks As Collection, prefix As String) As Boolean
    If pos_ > toks.Count Then
        Exit Function
    End If
    Dim v() As Variant
    v = toks(pos_)
    If v(0) = TK_PUNCT And v(1) = prefix Then
        Consume = True
        pos_ = pos_ + 1
    End If
End Function

Private Sub Expect(toks As Collection, prefix As String)
    If Not Consume(toks, prefix) Then
        Call ErrorAt2(toks, "expected " & "'" & prefix & "'")
    End If
End Sub

Private Sub ErrorAt2(toks As Collection, msg As String)
    Dim start As Long
    If pos_ <= toks.Count Then
        Dim t() As Variant
        t = toks(pos_)
        start = t(2)
    Else
        start = pos_
    End If
    Dim prefix As String
    prefix = String(4, " ")
    Debug.Print "parse error:"
    Debug.Print prefix & input_
    Debug.Print prefix & String(start - 1, " ") & "^ " & msg
    Debug.Print
    End
End Sub

' <expr>    ::= <equality>
Private Function Expr(toks As Collection) As Dictionary
    Set Expr = Equality(toks)
End Function

' <equality> ::= <relational> ("=" <relational> | "<>" <relational>)*
Private Function Equality(toks As Collection) As Dictionary
    Dim node As Dictionary
    Set node = Relational(toks)
    Do
        If Consume(toks, "=") Then
            Set node = NewBinary(ND_EQ, node, Relational(toks))
        ElseIf Consume(toks, "<>") Then
            Set node = NewBinary(ND_NE, node, Relational(toks))
        Else
            Set Equality = node
            Exit Function
        End If
    Loop
End Function

' <relational> ::= <add> ("<" <add> | "<=" <add> | ">" <add> | ">=" <add>)*
Private Function Relational(toks As Collection) As Dictionary
    Dim node As Dictionary
    Set node = Add(toks)
    Do
        If Consume(toks, "<") Then
            Set node = NewBinary(ND_LT, node, Add(toks))
        ElseIf Consume(toks, "<=") Then
            Set node = NewBinary(ND_LE, node, Add(toks))
        ElseIf Consume(toks, ">") Then
            Set node = NewBinary(ND_GT, node, Add(toks))
        ElseIf Consume(toks, ">=") Then
            Set node = NewBinary(ND_GE, node, Add(toks))
        Else
            Set Relational = node
            Exit Function
        End If
    Loop
End Function

' <add>    ::= <mul> ("+" <mul> | "-" <mul>)*
Private Function Add(toks As Collection) As Dictionary
    Dim node As Dictionary
    Set node = Mul(toks)
    Do
        If Consume(toks, "+") Then
            Set node = NewBinary(ND_ADD, node, Mul(toks))
        ElseIf Consume(toks, "-") Then
            Set node = NewBinary(ND_SUB, node, Mul(toks))
        Else
            Set Add = node
            Exit Function
        End If
    Loop
End Function

' <mul>     ::= <unary> ("*" <unary> | "/" <unary>)*
Private Function Mul(toks As Collection) As Dictionary
    Dim node As Dictionary
    Set node = Unary(toks)
    Do
        If Consume(toks, "*") Then
            Set node = NewBinary(ND_MUL, node, Unary(toks))
        ElseIf Consume(toks, "/") Then
            Set node = NewBinary(ND_DIV, node, Unary(toks))
        Else
            Set Mul = node
            Exit Function
        End If
    Loop
End Function

' <unary>   ::= ("+" | "-")? <primary>
Private Function Unary(toks As Collection) As Dictionary
    If Consume(toks, "+") Then
        Set Unary = Primary(toks)
    ElseIf Consume(toks, "-") Then
        Set Unary = NewBinary(ND_SUB, NewNum(0), Primary(toks))
    Else
        Set Unary = Primary(toks)
    End If
End Function

' <primary> ::= <num> | <ident> | <funcname> "(" <args>? ")" | "(" <expr> ")"
Private Function Primary(toks As Collection) As Dictionary
    If Consume(toks, "(") Then
        Dim node As Dictionary
        Set node = Expr(toks)
        Call Expect(toks, ")")
        node("enclosed") = True
        Set Primary = node
        Exit Function
    End If

    Dim t() As Variant
    t = toks(pos_)

    If t(0) = TK_NUM Then
        Set Primary = NewNum(CLng(t(1)))
        pos_ = pos_ + 1
        Exit Function
    End If

    If t(0) = TK_IDENT Then
        Set Primary = NewIdent(CStr(t(1)))
        pos_ = pos_ + 1
        Exit Function
    End If

    If t(0) = TK_FUNCNAME Then
        pos_ = pos_ + 1
        Call Expect(toks, "(")
        Dim args_ As Collection
        If Consume(toks, ")") Then
            Set args_ = New Collection
        Else
            Set args_ = Args(toks)
            Call Expect(toks, ")")
        End If
        Set Primary = NewFunc(CStr(t(1)), args_)
        Exit Function
    End If

    Call ErrorAt2(toks, "expected a number or an ident or an expression")
End Function

' <args> ::= <expr> ("," <expr>)*
Private Function Args(toks As Collection) As Collection
    Dim c As Collection
    Set c = New Collection
    c.Add Expr(toks)
    Do While Consume(toks, ",")
        c.Add Expr(toks)
    Loop

    Set Args = c
End Function

Public Function Pretty(node As Dictionary, indentLength As Long, Optional indentLevel As Long = 0) As String
    Dim buf As String
    Dim pos As Long
    Dim indent As String
    buf = String(256, vbNullChar)
    pos = 1
    indent = NewIndent(indentLevel, indentLength)
    Dim k As NodeKind
    Dim v As Variant
    k = node("kind")
    Select Case k
        Case ND_NUM, ND_IDENT
            Call PushString(buf, pos, node("val"))
        Case ND_ADD, ND_SUB, ND_MUL, ND_DIV, _
             ND_EQ, ND_NE, ND_LT, ND_LE, ND_GT, ND_GE
            If node("enclosed") Then
                Call PushString(buf, pos, "(")
                Call PushString(buf, pos, vbCrLf)
                Call PushString(buf, pos, indent)
                Call PushString(buf, pos, Pretty(node("lhs"), indentLength, indentLevel + 1))
                Call PushString(buf, pos, " ")
                Call PushString(buf, pos, OperatorMap(k))
                Call PushString(buf, pos, " ")
                Call PushString(buf, pos, Pretty(node("rhs"), indentLength, indentLevel + 1))
                Call PushString(buf, pos, vbCrLf)
                Call PushString(buf, pos, NewIndent(indentLevel - 1, indentLength))
                Call PushString(buf, pos, ")")
            Else
                Call PushString(buf, pos, Pretty(node("lhs"), indentLength, indentLevel + 1))
                Call PushString(buf, pos, " ")
                Call PushString(buf, pos, OperatorMap(k))
                Call PushString(buf, pos, " ")
                Call PushString(buf, pos, Pretty(node("rhs"), indentLength, indentLevel + 1))
            End If
        Case ND_FUNC
            Call PushString(buf, pos, node("name"))
            Call PushString(buf, pos, "(")
            Dim args_ As Collection
            Set args_ = node("args")
            If args_.Count = 0 Then
                Call PushString(buf, pos, ")")
            Else
                Call PushString(buf, pos, vbCrLf)
                Dim i As Long
                For i = 1 To args_.Count
                    Call PushString(buf, pos, NewIndent(indentLevel + 1, indentLength))
                    Call PushString(buf, pos, Pretty(args_(i), indentLength, indentLevel + 1))
                    If i < args_.Count Then
                        Call PushString(buf, pos, ",")
                        Call PushString(buf, pos, vbCrLf)
                    End If
                Next i
                Call PushString(buf, pos, vbCrLf)
                Call PushString(buf, pos, NewIndent(indentLevel, indentLength))
                Call PushString(buf, pos, ")")
            End If
        Case Else
    End Select

    Pretty = Mid(buf, 1, pos - 1)
End Function

Private Function NewIndent(level As Long, length As Long) As String
    If level <= 0 Then
        NewIndent = ""
        Exit Function
    End If
    NewIndent = Space(level * length)
End Function

Private Sub PushString(ByRef buf As String, start As Long, val As String)
    Do While (Len(buf) - start) + 1 < Len(val)
        Call DoubleBuffer(buf)
    Loop
    Mid(buf, start) = val
    start = start + Len(val)
End Sub

Private Sub DoubleBuffer(ByRef buf As String)
    Dim curLen As Long
    curLen = Len(buf)
    If curLen * 2 > BUF_MAX Then
        Debug.Print "error: The buffer has reached its maximum allowed size of " & BUF_MAX & " characters."
        End
    End If
    Dim newBuf As String
    newBuf = String(curLen * 2, vbNullChar)
    Mid(newBuf, 1) = buf
    buf = newBuf
End Sub
