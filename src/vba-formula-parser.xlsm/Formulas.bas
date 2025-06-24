Attribute VB_Name = "Formulas"
Option Explicit

Private input_ As String
Public Enum TokenKind
    TK_NUM
    TK_PUNCT
    TK_IDENT
    TK_FUNCNAME
    TK_STRING
End Enum

Private Type Token
    kind As TokenKind
    val As Variant
    pos As Long
End Type

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
    ND_STRING
    ND_CONCAT
    ND_ARRAY
End Enum

Private Type Parser
    tokens As Collection
    pos As Long
End Type

Private Type StringBuffer
    buf As String
    pos As Long
End Type

Private Const BUF_MAX As Long = 8096

Static Property Get TokenKindMap() As Dictionary
    Set TokenKindMap = New Dictionary
    TokenKindMap.Add TK_NUM, "TK_NUM"
    TokenKindMap.Add TK_PUNCT, "TK_PUNCT"
    TokenKindMap.Add TK_IDENT, "TK_IDENT"
    TokenKindMap.Add TK_FUNCNAME, "TK_FUNCNAME"
    TokenKindMap.Add TK_STRING, "TK_STRING"
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
    NodeKindMap.Add ND_STRING, "ND_STRING"
    NodeKindMap.Add ND_CONCAT, "ND_CONCAT"
    NodeKindMap.Add ND_ARRAY, "ND_ARRAY"
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
    OperatorMap.Add ND_CONCAT, "&"
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
            Case c = "(" Or c = ")" Or c = "{" Or c = "}"
                toks.Add NewToken(TK_PUNCT, c, i)
                i = i + 1
            Case c = ","
                toks.Add NewToken(TK_PUNCT, c, i)
                i = i + 1
            Case c = ";"
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
                        ErrorAt str, "expected '('"
                    End If
                    toks.Add NewToken(TK_FUNCNAME, Mid(str, start, i - start), start)
                Else
                    If IsNumeric(Mid(str, i - 1, 1)) Then
                        ErrorAt str, "expected a char"
                    End If
                    If Mid(str, i, 1) = "(" Then
                        toks.Add NewToken(TK_FUNCNAME, Mid(str, start, i - start), start)
                    Else
                        toks.Add NewToken(TK_IDENT, Mid(str, start, i - start), start)
                    End If
                End If
                expectFuncName = False
            Case c = """"
                start = i
                i = i + 1
                Do
                    If i > Len(str) Then
                        ErrorAt Mid(str, start), "unclosed string literal"
                    End If
                    If Mid(str, i, 1) = """" Then
                        Exit Do
                    End If
                    i = i + 1
                Loop
                ' the surrounding quotes are not needed.
                toks.Add NewToken(TK_STRING, Mid(str, start + 1, i - start - 1), start)
                i = i + 1
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
            Case c = "&"
                toks.Add NewToken(TK_PUNCT, "&", i)
                i = i + 1
            Case Else
                ErrorAt Mid(str, i), "unexpected token"
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
    IsIdent = (97 <= dec And dec <= 122) Or (65 <= dec And dec <= 90) Or c = "_" Or c = "\"
End Function

Private Sub ErrorAt(rest As String, msg As String)
    Dim pos As Long
    pos = Len(input_) - Len(rest)
    Dim prefix As String
    prefix = Space(4)
    Debug.Print "tokenize error:"
    Debug.Print prefix & input_
    Debug.Print prefix & Space(pos) & "^ " & msg
    Debug.Print
    End
End Sub

Public Function Parse(str As String) As Dictionary
    Dim p As Parser
    p = NewParser(Tokenize(str))

    Dim root As Dictionary
    Set root = Expr(p)
    If HasNext(p) Then
        ErrorAt2 p, "unexpected trailing token"
    End If

    Set Parse = root
End Function

Private Function NewParser(tokens As Collection) As Parser
    Dim p As Parser
    Set p.tokens = tokens
    p.pos = 1
    NewParser = p
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

Private Function NewString(val As String) As Dictionary
    Set NewString = NewNode(ND_STRING)
    NewString.Add "val", val
End Function

Private Function NewFunc(name_ As String, args_ As Collection) As Dictionary
    Set NewFunc = NewNode(ND_FUNC)
    NewFunc.Add "name", name_
    NewFunc.Add "args", args_
End Function

Private Function NewArray(elems As Collection) As Dictionary
    Set NewArray = NewNode(ND_ARRAY)
    NewArray.Add "elements", elems
End Function

Private Sub Advance(p As Parser)
    p.pos = p.pos + 1
End Sub

Private Sub Rewind(p As Parser)
    p.pos = p.pos - 1
End Sub

Private Function HasNext(p As Parser) As Boolean
    HasNext = (p.pos <= p.tokens.Count)
End Function

Private Function NextToken(p As Parser) As Token
    If Not HasNext(p) Then
        ErrorAt2 p, "no more tokens can be parsed"
    End If

    Dim t() As Variant
    t = p.tokens(p.pos)
    Advance p

    Dim tok As Token
    tok.kind = t(0)
    tok.val = t(1)
    tok.pos = t(2)

    NextToken = tok
End Function

Private Function Consume(p As Parser, prefix As String) As Boolean
    Dim t As Token
    t = NextToken(p)
    If t.kind = TK_PUNCT And t.val = prefix Then
        Consume = True
    Else
        Rewind p
    End If
End Function

Private Sub Expect(p As Parser, prefix As String)
    If Not Consume(p, prefix) Then
        ErrorAt2 p, "expected " & "'" & prefix & "'"
    End If
End Sub

Private Sub ErrorAt2(p As Parser, msg As String)
    Dim start As Long
    If HasNext(p) Then
        Dim t As Token
        t = NextToken(p)
        start = t.pos
    Else
        start = p.pos
    End If
    Dim prefix As String
    prefix = Space(4)
    Debug.Print "parse error:"
    Debug.Print prefix & input_
    Debug.Print prefix & Space(start - 1) & "^ " & msg
    Debug.Print
    End
End Sub

' <expr>    ::= <equality>
Private Function Expr(p As Parser) As Dictionary
    Set Expr = Equality(p)
End Function

' <equality> ::= <relational> ("=" <relational> | "<>" <relational>)*
Private Function Equality(p As Parser) As Dictionary
    Dim node As Dictionary
    Set node = Relational(p)
    Do While HasNext(p)
        If Consume(p, "=") Then
            Set node = NewBinary(ND_EQ, node, Relational(p))
        ElseIf Consume(p, "<>") Then
            Set node = NewBinary(ND_NE, node, Relational(p))
        Else
            Exit Do
        End If
    Loop
    Set Equality = node
End Function

' <relational> ::= <concatenation> (("<" | "<=" | ">" | ">=") <concatenation>)*
Private Function Relational(p As Parser) As Dictionary
    Dim node As Dictionary
    Set node = Concatenation(p)
    Do While HasNext(p)
        If Consume(p, "<") Then
            Set node = NewBinary(ND_LT, node, Add(p))
        ElseIf Consume(p, "<=") Then
            Set node = NewBinary(ND_LE, node, Add(p))
        ElseIf Consume(p, ">") Then
            Set node = NewBinary(ND_GT, node, Add(p))
        ElseIf Consume(p, ">=") Then
            Set node = NewBinary(ND_GE, node, Add(p))
        Else
            Exit Do
        End If
    Loop
    Set Relational = node
End Function

' <concatenation> ::= <add> ("&" <add>)*
Private Function Concatenation(p As Parser) As Dictionary
    Dim node As Dictionary
    Set node = Add(p)
    Do While HasNext(p)
        If Consume(p, "&") Then
            Set node = NewBinary(ND_CONCAT, node, Add(p))
        Else
            Exit Do
        End If
    Loop
    Set Concatenation = node
End Function

' <add>    ::= <mul> ("+" <mul> | "-" <mul>)*
Private Function Add(p As Parser) As Dictionary
    Dim node As Dictionary
    Set node = Mul(p)
    Do While HasNext(p)
        If Consume(p, "+") Then
            Set node = NewBinary(ND_ADD, node, Mul(p))
        ElseIf Consume(p, "-") Then
            Set node = NewBinary(ND_SUB, node, Mul(p))
        Else
            Exit Do
        End If
    Loop
    Set Add = node
End Function

' <mul>     ::= <unary> ("*" <unary> | "/" <unary>)*
Private Function Mul(p As Parser) As Dictionary
    Dim node As Dictionary
    Set node = Unary(p)
    Do While HasNext(p)
        If Consume(p, "*") Then
            Set node = NewBinary(ND_MUL, node, Unary(p))
        ElseIf Consume(p, "/") Then
            Set node = NewBinary(ND_DIV, node, Unary(p))
        Else
            Exit Do
        End If
    Loop
    Set Mul = node
End Function

' <unary>   ::= ("+" | "-")? <primary>
Private Function Unary(p As Parser) As Dictionary
    If Consume(p, "+") Then
        Set Unary = Primary(p)
    ElseIf Consume(p, "-") Then
        Set Unary = NewBinary(ND_SUB, NewNum(0), Primary(p))
    Else
        Set Unary = Primary(p)
    End If
End Function

' <primary> ::= <num> | <ident> | <string> | "{" <constants> "}" | <funcname> "(" <args>? ")" | "(" <expr> ")"
Private Function Primary(p As Parser) As Dictionary
    If Consume(p, "(") Then
        Dim node As Dictionary
        Set node = Expr(p)
        Expect p, ")"
        node("enclosed") = True
        Set Primary = node
        Exit Function
    End If

    If Consume(p, "{") Then
        Dim elems As Collection
        Set elems = Constants(p)
        Expect p, "}"
        Set Primary = NewArray(elems)
        Exit Function
    End If

    Dim t As Token
    t = NextToken(p)

    If t.kind = TK_NUM Then
        Set Primary = NewNum(CLng(t.val))
        Exit Function
    End If

    If t.kind = TK_IDENT Then
        Set Primary = NewIdent(CStr(t.val))
        Exit Function
    End If

    If t.kind = TK_STRING Then
        Set Primary = NewString(CStr(t.val))
        Exit Function
    End If

    If t.kind = TK_FUNCNAME Then
        Expect p, "("
        Dim args_ As Collection
        If Consume(p, ")") Then
            Set args_ = New Collection
        Else
            Set args_ = Args(p)
            Expect p, ")"
        End If
        Set Primary = NewFunc(CStr(t.val), args_)
        Exit Function
    End If

    ErrorAt2 p, "expected a number or an ident or an expression"
End Function

' <constants> ::= <constant> (("," | ";") <constant>)*
Private Function Constants(p As Parser) As Collection
    Dim c As Collection
    Set c = New Collection
    c.Add Constant(p)
    Do While HasNext(p)
        If Consume(p, ",") Then
            c.Add Constant(p)
        ElseIf Consume(p, ";") Then
            c.Add Constant(p)
        Else
            Exit Do
        End If
    Loop
    Set Constants = c
End Function

' <constant> ::= <num> | <string> | "TRUE" | "FALSE"
Private Function Constant(p As Parser) As Dictionary
    Dim t As Token
    t = NextToken(p)

    If t.kind = TK_NUM Then
        Set Constant = NewNum(CLng(t.val))
        Exit Function
    End If

    If t.kind = TK_STRING Then
        Set Constant = NewString(CStr(t.val))
        Exit Function
    End If

    If t.kind = TK_IDENT And (t.val = "TRUE" Or t.val = "FALSE") Then
        Set Constant = NewIdent(CStr(t.val))
        Exit Function
    End If

    ErrorAt2 p, "expected a costant value"
End Function

' <args> ::= <expr> ("," <expr>)*
Private Function Args(p As Parser) As Collection
    Dim c As Collection
    Set c = New Collection
    c.Add Expr(p)
    Do While Consume(p, ",")
        c.Add Expr(p)
    Loop

    Set Args = c
End Function

Public Function Pretty(node As Dictionary, indentLength As Long, Optional indentLevel As Long = 0) As String
    Dim sb As StringBuffer
    Dim k As NodeKind
    sb = NewStringBuffer(256)
    k = node("kind")
    Select Case k
        Case ND_NUM, ND_IDENT
            Push sb, node("val")
        Case ND_STRING
            Push sb, Chr(34)
            Push sb, node("val")
            Push sb, Chr(34)
        Case ND_ADD, ND_SUB, ND_MUL, ND_DIV, _
             ND_EQ, ND_NE, ND_LT, ND_LE, ND_GT, ND_GE, _
             ND_CONCAT
            If node("enclosed") Then
                Push sb, "("
                Push sb, vbCrLf
                Push sb, NewIndent(indentLevel + 1, indentLength)
                Push sb, Pretty(node("lhs"), indentLength, indentLevel + 1)
                Push sb, " "
                Push sb, OperatorMap(k)
                Push sb, " "
                Push sb, Pretty(node("rhs"), indentLength, indentLevel + 1)
                Push sb, vbCrLf
                Push sb, NewIndent(indentLevel - 1, indentLength)
                Push sb, ")"
            Else
                Push sb, Pretty(node("lhs"), indentLength, indentLevel)
                Push sb, " "
                Push sb, OperatorMap(k)
                Push sb, " "
                Push sb, Pretty(node("rhs"), indentLength, indentLevel)
            End If
        Case ND_FUNC
            Push sb, node("name")
            Push sb, "("
            Dim args_ As Collection
            Set args_ = node("args")
            If args_.Count = 0 Then
                Push sb, ")"
            Else
                Push sb, vbCrLf
                Dim i As Long
                For i = 1 To args_.Count
                    Push sb, NewIndent(indentLevel + 1, indentLength)
                    Push sb, Pretty(args_(i), indentLength, indentLevel + 1)
                    If i < args_.Count Then
                        Push sb, ","
                        Push sb, vbCrLf
                    End If
                Next i
                Push sb, vbCrLf
                Push sb, NewIndent(indentLevel, indentLength)
                Push sb, ")"
            End If
        Case ND_STRING
            Push sb, Chr(34)
            Push sb, node("val")
            Push sb, Chr(34)
        Case Else
    End Select

    Pretty = ToString(sb)
End Function

Private Function NewIndent(level As Long, length As Long) As String
    If level <= 0 Then
        NewIndent = ""
        Exit Function
    End If
    NewIndent = Space(level * length)
End Function

Public Function NewStringBuffer(size As Long) As StringBuffer
    If size > BUF_MAX Then
        Debug.Print "Error: Requested buffer size (" & size & ") exceeds maximum allowed (" & BUF_MAX & " characters)."
        End
    End If
    Dim sb As StringBuffer
    sb.buf = String(size, vbNullChar)
    sb.pos = 1
    NewStringBuffer = sb
End Function

Public Sub Push(sb As StringBuffer, val As String)
    Do While Len(val) > (Len(sb.buf) - sb.pos) + 1
        DoubleBuffer sb
    Loop
    Mid(sb.buf, sb.pos) = val
    sb.pos = sb.pos + Len(val)
End Sub

Private Sub DoubleBuffer(sb As StringBuffer)
    Dim curLen As Long
    curLen = Len(sb.buf)
    If curLen * 2 > BUF_MAX Then
        Debug.Print "error: The buffer has reached its maximum allowed size of " & BUF_MAX & " characters."
        End
    End If
    Dim newBuf As String
    newBuf = String(curLen * 2, vbNullChar)
    Mid(newBuf, 1) = sb.buf
    sb.buf = newBuf
End Sub

Public Function ToString(sb As StringBuffer) As String
    ToString = Mid(sb.buf, 1, sb.pos - 1)
End Function
