Attribute VB_Name = "Formulas"
Option Explicit

Private input_ As String
Private Type Tokenizer
    input As String
    pos As Long
    mark As Long
End Type

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
    ND_ARRAY_ROW
    ND_EMPTY
End Enum

Private Type Parser
    tokens As Collection
    pos As Long
End Type

Private Type StringBuffer
    buf As String
    pos As Long
End Type

Public Type Indentation
    char As String
    length As Long
    level As Long
End Type

Public Type Formatter
    indent As Indentation
    newLine As String
    eqAtStart As Boolean
    newLineAtEof As Boolean
End Type

Private Const BUF_MAX As Long = 16384

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
    NodeKindMap.Add ND_ARRAY_ROW, "ND_ARRAY_ROW"
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
    input_ = Replace(Replace(str, vbCr, " "), vbLf, " ")
    Dim t As Tokenizer
    t.input = input_
    t.pos = 1
    t.mark = 0

    Dim toks As Collection
    Set toks = New Collection
    Do While Tokenizer_HasNext(t)
        toks.Add Tokenizer_NextToken(t)
    Loop

    Set Tokenize = toks
End Function

Private Sub Tokenizer_Mark(t As Tokenizer)
    t.mark = t.pos
End Sub

Private Function Tokenizer_Capture(t As Tokenizer) As String
    Tokenizer_Capture = Mid(t.input, t.mark, t.pos - t.mark)
End Function

Private Function Tokenizer_HasNext(t As Tokenizer) As Boolean
    Tokenizer_HasNext = (t.pos <= Len(t.input))
End Function

Private Function Tokenizer_Consume(t As Tokenizer, prefix As String) As Boolean
    Dim n As Long
    n = Len(prefix)
    If Mid(t.input, t.pos, n) = prefix Then
        t.pos = t.pos + n
        Tokenizer_Consume = True
        Exit Function
    End If
    Tokenizer_Consume = False
End Function

Private Function Tokenizer_ConsumeAny(t As Tokenizer, ParamArray prefixes() As Variant) As Boolean
    Dim prefix As Variant
    For Each prefix In prefixes
        If Tokenizer_Consume(t, CStr(prefix)) Then
            Tokenizer_ConsumeAny = True
            Exit Function
        End If
    Next prefix
    Tokenizer_ConsumeAny = False
End Function

Private Sub Tokenizer_SkipWhitespaces(t As Tokenizer)
    Do While Tokenizer_ConsumeAny(t, " ", vbCrLf, vbLf)
    Loop
End Sub

Private Function Tokenizer_NextToken(t As Tokenizer) As Variant()
    Tokenizer_SkipWhitespaces t
    Dim c As String
    c = Mid(t.input, t.pos, 1)
    Select Case True
        Case IsNumeric(c)
            Tokenizer_Mark t
            Do
                t.pos = t.pos + 1
            Loop While IsNumeric(Mid(t.input, t.pos, 1))
            Tokenizer_NextToken = NewToken(TK_NUM, Tokenizer_Capture(t), t.mark)
        Case c = "+" Or c = "-" Or c = "*" Or c = "/"
            Tokenizer_NextToken = NewToken(TK_PUNCT, c, t.pos)
            t.pos = t.pos + 1
        Case c = "(" Or c = ")" Or c = "{" Or c = "}"
            Tokenizer_NextToken = NewToken(TK_PUNCT, c, t.pos)
            t.pos = t.pos + 1
        Case c = ","
            Tokenizer_NextToken = NewToken(TK_PUNCT, c, t.pos)
            t.pos = t.pos + 1
        Case c = ";"
            Tokenizer_NextToken = NewToken(TK_PUNCT, c, t.pos)
            t.pos = t.pos + 1
        Case c = "."
            Tokenizer_NextToken = NewToken(TK_PUNCT, c, t.pos)
            t.pos = t.pos + 1
        Case IsIdent(c)
            Tokenizer_Mark t
            Do
                t.pos = t.pos + 1
            Loop While IsIdent(Mid(t.input, t.pos, 1)) Or IsNumeric(Mid(t.input, t.pos, 1))
            If Mid(t.input, t.pos, 1) = "(" Then
                Tokenizer_NextToken = NewToken(TK_FUNCNAME, Tokenizer_Capture(t), t.mark)
            Else
                Tokenizer_NextToken = NewToken(TK_IDENT, Tokenizer_Capture(t), t.mark)
            End If
        Case c = """"
            Tokenizer_Mark t
            Do
                t.pos = t.pos + 1
                If Not Tokenizer_HasNext(t) Then
                    ErrorAt t, "unclosed string literal"
                End If
                If Mid(t.input, t.pos, 1) = """" Then
                    t.pos = t.pos + 1
                    Exit Do
                End If
            Loop
            Tokenizer_NextToken = NewToken(TK_STRING, Tokenizer_Capture(t), t.mark)
        Case c = "="
            Tokenizer_NextToken = NewToken(TK_PUNCT, c, t.pos)
            t.pos = t.pos + 1
        Case c = "<"
            If Mid(t.input, t.pos + 1, 1) = ">" Or Mid(t.input, t.pos + 1, 1) = "=" Then
                Tokenizer_NextToken = NewToken(TK_PUNCT, Mid(t.input, t.pos, 2), t.pos)
                t.pos = t.pos + 2
            Else
                Tokenizer_NextToken = NewToken(TK_PUNCT, c, t.pos)
                t.pos = t.pos + 1
            End If
        Case c = ">"
            If Mid(t.input, t.pos + 1, 1) = "=" Then
                Tokenizer_NextToken = NewToken(TK_PUNCT, Mid(t.input, t.pos, 2), t.pos)
                t.pos = t.pos + 2
            Else
                Tokenizer_NextToken = NewToken(TK_PUNCT, c, t.pos)
                t.pos = t.pos + 1
            End If
        Case c = "&"
            Tokenizer_NextToken = NewToken(TK_PUNCT, "&", t.pos)
            t.pos = t.pos + 1
        Case Else
            ErrorAt t, "unexpected token"
    End Select
End Function

Private Function NewToken(kind As Long, val As String, col As Long) As Variant()
    NewToken = Array(kind, val, col)
End Function

Private Function IsWhitespace(c As String) As Boolean
    IsWhitespace = (c = " ") Or (c = vbCr) Or (c = vbLf)
End Function

Private Function IsIdent(c As String) As Boolean
    If c = "" Then
        IsIdent = False
        Exit Function
    End If
    Dim dec As Long
    dec = Asc(c)
    IsIdent = (97 <= dec And dec <= 122) Or (65 <= dec And dec <= 90) Or c = "_" Or c = "\" Or c = "."
End Function

Private Sub ErrorAt(t As Tokenizer, msg As String)
    Dim prefix As String
    Dim m As String
    prefix = IndentString(NewIndentation(" ", 1, 4))
    m = m & "tokenize error:" & vbCrLf
    m = m & prefix & t.input & vbCrLf
    m = m & prefix & IndentString(NewIndentation(" ", 1, t.pos - 1)) & "^ " & msg
    Err.Raise 5, Description:=m
End Sub

Public Function Parse(str As String) As Dictionary
    Dim p As Parser
    p = NewParser(Tokenize(str))
    If Not Consume(p, "=") Then
        ErrorAt2 p, "expected '='"
    End If

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

Private Function NewArrayRow(elems As Collection) As Dictionary
    Set NewArrayRow = NewNode(ND_ARRAY_ROW)
    NewArrayRow.Add "elements", elems
End Function

Private Function NewEmpty() As Dictionary
    Set NewEmpty = NewNode(ND_EMPTY)
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

Private Function Peek(p As Parser) As Token
    Dim t As Token
    t = NextToken(p)
    Rewind p
    Peek = t
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
        start = t.pos - 1
    Else
        start = p.pos
    End If
    Dim prefix As String
    Dim m As String
    prefix = IndentString(NewIndentation(" ", 1, 4))
    m = m & "parse error:" & vbCrLf
    m = m & prefix & input_ & vbCrLf
    m = m & prefix & IndentString(NewIndentation(" ", 1, start - 1)) & "^ " & msg
    Err.Raise 5, Description:=m
End Sub

' <expr>    ::= "=" <equality>
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

' <primary> ::= <num> | <ident> | <string> | "{" <array_rows> "}" | <funcname> "(" <args>? ")" | "(" <expr> ")"
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
        Dim r As Collection
        Set r = ArrayRows(p)
        Expect p, "}"
        Set Primary = NewArray(r)
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

' <array_rows> ::= <array_row> (";" <array_row>)*
Private Function ArrayRows(p As Parser) As Collection
    Dim c As Collection
    Set c = New Collection
    c.Add ArrayRow(p)
    Do While HasNext(p)
        If Consume(p, ";") Then
            c.Add ArrayRow(p)
        Else
            Exit Do
        End If
    Loop
    Set ArrayRows = c
End Function

' <array_row> ::= <constant> (","  <constant>)*
Private Function ArrayRow(p As Parser) As Dictionary
    Dim c As Collection
    Set c = New Collection
    c.Add Constant(p)
    Do While HasNext(p)
        If Consume(p, ",") Then
            c.Add Constant(p)
        Else
            Exit Do
        End If
    Loop
    Set ArrayRow = NewArrayRow(c)
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

' <args> ::= <expr> ("," <expr>?)*
Private Function Args(p As Parser) As Collection
    Dim c As Collection
    Set c = New Collection
    c.Add Expr(p)
    Do While Consume(p, ",")
        Dim t As Token
        t = Peek(p)
        If t.kind = TK_PUNCT And t.val = "," Then
            c.Add NewEmpty()
        ElseIf t.kind = TK_PUNCT And t.val = ")" Then
            c.Add NewEmpty()
            Exit Do
        Else
            c.Add Expr(p)
        End If
    Loop

    Set Args = c
End Function

Public Function Stringify(ast As Dictionary, fmt As Formatter) As String
    Dim s As String
    s = Pretty(ast, fmt)
    If fmt.eqAtStart Then
        s = "=" & s
    End If
    If fmt.newLineAtEof Then
        s = s & vbCrLf
    End If
    Stringify = s
End Function

Public Function DebugAst(ast As Dictionary, fmt As Formatter) As String
    Dim json As String
    json = ToJson(ast, fmt)
    DebugAst = json
End Function

Public Function NewIndentation(char As String, length As Long, Optional level As Long = 0) As Indentation
    Dim indent As Indentation
    indent.char = char
    indent.level = level
    indent.length = length
    NewIndentation = indent
End Function

Public Function NewFormatter( _
    indent As Indentation, _
    newLine As String, _
    eqAtStart As Boolean, _
    newLineAtEof As Boolean) As Formatter
    Dim f As Formatter
    f.indent = indent
    f.newLine = newLine
    f.eqAtStart = eqAtStart
    f.newLineAtEof = newLineAtEof
    NewFormatter = f
End Function

Public Property Get DefaultFormatter() As Formatter
    DefaultFormatter = NewFormatter( _
        NewIndentation("", 0), _
        "", _
        True, _
        True _
    )
End Property

Private Function Pretty(node As Dictionary, fmt As Formatter) As String
    Dim sb As StringBuffer
    Dim k As NodeKind
    sb = NewStringBuffer(256)
    k = node("kind")
    Dim i As Long
    Select Case k
        Case ND_NUM, ND_IDENT, ND_STRING
            Push sb, node("val")
        Case ND_ADD, ND_SUB, ND_MUL, ND_DIV, _
             ND_EQ, ND_NE, ND_LT, ND_LE, ND_GT, ND_GE, _
             ND_CONCAT
            If node("enclosed") Then
                Push sb, "("
                Push sb, fmt.newLine
                Push sb, NextIndent(fmt)
                Push sb, Pretty(node("lhs"), UpIndent(fmt))
                Push sb, " "
                Push sb, OperatorMap(k)
                Push sb, " "
                Push sb, Pretty(node("rhs"), UpIndent(fmt))
                Push sb, fmt.newLine
                Push sb, PrevIndent(fmt)
                Push sb, ")"
            Else
                Push sb, Pretty(node("lhs"), fmt)
                Push sb, " "
                Push sb, OperatorMap(k)
                Push sb, " "
                Push sb, Pretty(node("rhs"), fmt)
            End If
        Case ND_FUNC
            Push sb, node("name")
            Push sb, "("
            Dim args_ As Collection
            Set args_ = node("args")
            If args_.Count = 0 Then
                Push sb, ")"
            Else
                Push sb, fmt.newLine
                For i = 1 To args_.Count
                    Push sb, NextIndent(fmt)
                    Push sb, Pretty(args_(i), UpIndent(fmt))
                    If i < args_.Count Then
                        Push sb, ","
                        Push sb, fmt.newLine
                    End If
                Next i
                Push sb, fmt.newLine
                Push sb, CurrentIndent(fmt)
                Push sb, ")"
            End If
        Case ND_ARRAY
            Push sb, "{"
            Push sb, fmt.newLine
            Dim rows_ As Collection
            Set rows_ = node("elements")
            For i = 1 To rows_.Count
                Push sb, NextIndent(fmt)
                Push sb, Pretty(rows_(i), fmt)
                If i < rows_.Count Then
                    Push sb, ";"
                    Push sb, fmt.newLine
                End If
            Next i
            Push sb, fmt.newLine
            Push sb, CurrentIndent(fmt)
            Push sb, "}"
        Case ND_ARRAY_ROW
            Dim cols As Collection
            Set cols = node("elements")
            For i = 1 To cols.Count
                Push sb, Pretty(cols(i), fmt)
                If i < cols.Count Then
                    Push sb, ","
                    Push sb, " "
                End If
            Next i
        Case ND_EMPTY
            Push sb, ""
        Case Else
    End Select

    Pretty = ToString(sb)
End Function

Private Function ToJson(ast As Dictionary, fmt As Formatter) As String
    Dim i As Long
    Dim sb As StringBuffer
    sb = NewStringBuffer(256)
    Push sb, "{"
    Push sb, fmt.newLine
    Push sb, NextIndent(fmt)
    Push sb, """kind"": "
    Push sb, Quote(NodeKindMap(CLng(ast("kind"))))
    Push sb, ","
    Push sb, fmt.newLine
    Push sb, NextIndent(fmt)
    Select Case ast("kind")
        Case ND_NUM, ND_STRING
            Push sb, """val"": "
            Push sb, ast("val")
        Case ND_IDENT
            Push sb, """val"": "
            Push sb, Quote(ast("val"))
        Case ND_ADD, ND_SUB, ND_MUL, ND_DIV, _
             ND_EQ, ND_NE, _
             ND_LT, ND_LE, ND_GT, ND_GE, _
             ND_CONCAT
            Push sb, """lhs"": "
            Push sb, ToJson(ast("lhs"), UpIndent(fmt))
            Push sb, ","
            Push sb, fmt.newLine
            Push sb, NextIndent(fmt)
            Push sb, """rhs"": "
            Push sb, ToJson(ast("rhs"), UpIndent(fmt))
        Case ND_FUNC
            Push sb, """name"": "
            Push sb, Quote(ast("name"))
            Push sb, ","
            Push sb, fmt.newLine
            Push sb, NextIndent(fmt)
            Push sb, """args"": ["
            Dim args_ As Collection
            Set args_ = ast("args")
            If args_.Count > 0 Then
                Push sb, fmt.newLine
                For i = 1 To args_.Count
                    ' one level deeper than the array property itself
                    Push sb, NextIndent(UpIndent(fmt))
                    Push sb, ToJson(args_(i), UpIndent(UpIndent(fmt)))
                    If i < args_.Count Then
                        Push sb, ","
                    End If
                    Push sb, fmt.newLine
                Next i
                ' for closing bracket
                Push sb, NextIndent(fmt)
            End If
            Push sb, "]"
        Case ND_ARRAY
            Push sb, """elements"": ["
            Dim rows_ As Collection
            Set rows_ = ast("elements")
            If rows_.Count > 0 Then
                Push sb, fmt.newLine
                For i = 1 To rows_.Count
                    Push sb, NextIndent(UpIndent(fmt))
                    Push sb, ToJson(rows_(i), UpIndent(UpIndent(fmt)))
                    If i < rows_.Count Then
                        Push sb, ","
                    End If
                    Push sb, fmt.newLine
                Next i
                Push sb, NextIndent(fmt)
            End If
            Push sb, "]"
        Case ND_ARRAY_ROW
            Push sb, """elements"": ["
            Dim cols_ As Collection
            Set cols_ = ast("elements")
            If cols_.Count > 0 Then
                Push sb, fmt.newLine
                For i = 1 To cols_.Count
                    Push sb, NextIndent(UpIndent(fmt))
                    Push sb, ToJson(cols_(i), UpIndent(UpIndent(fmt)))
                    If i < cols_.Count Then
                        Push sb, ","
                    End If
                    Push sb, fmt.newLine
                Next i
                Push sb, NextIndent(fmt)
            End If
            Push sb, "]"
    End Select
    Push sb, fmt.newLine
    Push sb, CurrentIndent(fmt)
    Push sb, "}"

    ToJson = ToString(sb)
End Function

Private Function Quote(str As String) As String
    Quote = """" & str & """"
End Function

Private Function HasValue(indent As Indentation) As Boolean
    HasValue = indent.char <> "" And indent.level > 0 And indent.length > 0
End Function

Private Function IndentString(indent As Indentation) As String
    If Not HasValue(indent) Then
        IndentString = ""
        Exit Function
    End If
    IndentString = String(indent.length * indent.level, indent.char)
End Function

Private Function UpIndent(fmt As Formatter) As Formatter
    Dim newFmt As Formatter
    newFmt = NewFormatter( _
        NewIndentation(fmt.indent.char, fmt.indent.length, fmt.indent.level + 1), _
        fmt.newLine, _
        fmt.eqAtStart, _
        fmt.newLineAtEof _
    )
    UpIndent = newFmt
End Function

Private Function DownIndent(fmt As Formatter) As Formatter
    Dim newFmt As Formatter
    newFmt = NewFormatter( _
        NewIndentation(fmt.indent.char, fmt.indent.length, fmt.indent.level - 1), _
        fmt.newLine, _
        fmt.eqAtStart, _
        fmt.newLineAtEof _
    )
    DownIndent = newFmt
End Function

Private Function CurrentIndent(fmt As Formatter) As String
    CurrentIndent = IndentString(fmt.indent)
End Function

Private Function NextIndent(fmt As Formatter) As String
    NextIndent = CurrentIndent(UpIndent(fmt))
End Function

Private Function PrevIndent(fmt As Formatter) As String
    PrevIndent = CurrentIndent(DownIndent(fmt))
End Function

Private Function NewStringBuffer(size As Long) As StringBuffer
    If size > BUF_MAX Then
        Debug.Print "Error: Requested buffer size (" & size & ") exceeds maximum allowed (" & BUF_MAX & " characters)."
        End
    End If
    Dim sb As StringBuffer
    sb.buf = String(size, vbNullChar)
    sb.pos = 1
    NewStringBuffer = sb
End Function

Private Sub Push(sb As StringBuffer, val As String)
    If Len(val) = 0 Then
        Exit Sub
    End If
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

Private Function ToString(sb As StringBuffer) As String
    ToString = Mid(sb.buf, 1, sb.pos - 1)
End Function
