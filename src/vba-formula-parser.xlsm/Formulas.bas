Attribute VB_Name = "Formulas"
Option Explicit

Private input_ As String
Private Enum TokenKind
    TK_NUM
    TK_PUNCT
    TK_IDENT
End Enum

Private pos_ As Long
Private Enum NodeKind
    ND_NUM
    ND_ADD
    ND_SUB
    ND_MUL
    ND_DIV
    ND_IDENT
End Enum

Static Property Get TokenKindMap() As Dictionary
    Set TokenKindMap = New Dictionary
    TokenKindMap.Add TK_NUM, "TK_NUM"
    TokenKindMap.Add TK_PUNCT, "TK_PUNCT"
    TokenKindMap.Add TK_IDENT, "TK_IDENT"
End Property

Static Property Get NodeKindMap() As Dictionary
    Set NodeKindMap = New Dictionary
    NodeKindMap.Add ND_NUM, "ND_NUM"
    NodeKindMap.Add ND_ADD, "ND_ADD"
    NodeKindMap.Add ND_SUB, "ND_SUB"
    NodeKindMap.Add ND_MUL, "ND_MUL"
    NodeKindMap.Add ND_DIV, "ND_DIV"
    NodeKindMap.Add ND_IDENT, "ND_IDENT"
End Property

Private Function Tokenize(str As String) As Collection
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
                start = i
                Do
                    i = i + 1
                Loop While IsIdent(Mid(str, i, 1))
                toks.Add NewToken(TK_IDENT, Mid(str, start, i - start), start)
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

Public Function Parse(str As String) As Dictionary
    Dim toks As Collection
    Set toks = Tokenize(str)
    pos_ = 1

    Set Parse = Expr(toks)
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

Private Sub ErrorAt(rest As String, msg As String)
    Dim pos As Long
    pos = Len(input_) - Len(rest)
    Dim prefix As String
    prefix = String(4, " ")
    Debug.Print "error:"
    Debug.Print prefix & input_
    Debug.Print prefix & String(pos, " ") & "^ " & msg
    Debug.Print
    End
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
    Debug.Print "error:"
    Debug.Print prefix & input_
    Debug.Print prefix & String(start - 1, " ") & "^ " & msg
    Debug.Print
    End
End Sub

' <expr>    ::= <mul> ("+" <mul> | "-" <mul>)*
Private Function Expr(toks As Collection) As Dictionary
    Dim node As Dictionary
    Set node = Mul(toks)
    Do
        If Consume(toks, "+") Then
            Set node = NewBinary(ND_ADD, node, Mul(toks))
        ElseIf Consume(toks, "-") Then
            Set node = NewBinary(ND_SUB, node, Mul(toks))
        Else
            Set Expr = node
            Exit Function
        End If
    Loop
End Function

' <mul>     ::= <primary> ("*" <primary> | "/" <primary>)*
Private Function Mul(toks As Collection) As Dictionary
    Dim node As Dictionary
    Set node = Primary(toks)
    Do
        If Consume(toks, "*") Then
            Set node = NewBinary(ND_MUL, node, Primary(toks))
        ElseIf Consume(toks, "/") Then
            Set node = NewBinary(ND_DIV, node, Primary(toks))
        Else
            Set Mul = node
            Exit Function
        End If
    Loop
End Function

' <primary> ::= <num> | <ident> | "(" <expr> ")"
Private Function Primary(toks As Collection) As Dictionary
    If Consume(toks, "(") Then
        Dim node As Dictionary
        Set node = Expr(toks)
        Call Expect(toks, ")")
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

    Call ErrorAt2(toks, "expected a number or an ident or an expression")
End Function

Private Function ExpectNumber(toks As Collection) As Dictionary
    Dim t() As Variant
    t = toks(pos_)
    If t(0) <> TK_NUM Then
        Call ErrorAt2(toks, "expected a number")
    End If
    pos_ = pos_ + 1
    Set ExpectNumber = NewNum(CLng(t(1)))
End Function

Public Function Pretty(node As Dictionary) As String

End Function

Sub TestTokenize()
    Dim tests As Variant
    tests = Array( _
        "1+2", _
        "1+23*4/5", _
        "(1-23)*4", _
        "SUM(12)*3", _
        "SUM(12, 34)*5", _
        "" _
    )
    Dim t As Variant
    For Each t In tests
        If CStr(t) <> "" Then
            Debug.Print t
            Call DumpTokens(Tokenize(CStr(t)))
        End If
    Next t
End Sub

Sub TestParse()
    Dim tests As Variant
    tests = Array( _
        "1+2", _
        "1+2*3", _
        "(1+2)*3", _
        "x+y*z", _
        "(ab+cd)*ef", _
        "" _
    )
    Dim t As Variant
    For Each t In tests
        If CStr(t) <> "" Then
            Debug.Print t
            Call DumpNode(Parse(CStr(t)), 0)
        End If
    Next t
End Sub

Private Sub DumpTokens(toks As Collection)
    Dim t As Variant
    For Each t In toks
        Debug.Print "kind: " & TokenKindMap(t(0)) & ", val: " & t(1)
    Next t
    Debug.Print
End Sub

Private Sub DumpNode(node As Dictionary, indentLevel As Long)
    Dim k As NodeKind
    Dim indent As String
    Dim prefix As String
    k = node("kind")
    indent = String(indentLevel * 2, " ")
    prefix = indentLevel & " " & indent
    Select Case k
        Case ND_NUM, ND_IDENT
            Debug.Print prefix & "kind: " & NodeKindMap(k)
            Debug.Print prefix & "val: " & node("val")
        Case ND_ADD, ND_SUB, ND_MUL, ND_DIV
            Debug.Print prefix & "kind: " & NodeKindMap(k)
            Debug.Print prefix & "lhs:"
            Call DumpNode(node("lhs"), indentLevel + 1)
            Debug.Print prefix & "rhs:"
            Call DumpNode(node("rhs"), indentLevel + 1)
    End Select
    If indentLevel = 0 Then
        Debug.Print
    End If
End Sub

Private Sub AssertEq(x As Variant, y As Variant)
    If CStr(x) <> CStr(y) Then
        Debug.Print "assert failed left == " & x & ", right == " & y
    End If
End Sub
