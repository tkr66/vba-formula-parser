Attribute VB_Name = "FormulasTest"
Option Explicit

Sub TestTokenize()
    Dim tests As Variant
    tests = Array( _
        "1+2", _
        "1+23*4/5", _
        "(1-23)*4", _
        "SUM(12)*3", _
        "SUM(12, 34)*5", _
        "1=2<>3<4<=5>6>=7", _
        "" _
    )
    Dim t As Variant
    For Each t In tests
        If CStr(t) <> "" Then
            Debug.Print t
            Call DumpTokens(Formulas.Tokenize(CStr(t)))
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
        "+12*-3/+xyz", _
        "1=2<>3<4<=5>6>=7", _
        "(((((1=2)<>3)<4)<=5)>6)>=7", _
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
