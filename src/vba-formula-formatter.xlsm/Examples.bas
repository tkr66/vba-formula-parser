Attribute VB_Name = "Examples"
Option Explicit

Private Type Examples
    val As Collection
End Type

Private Type Example
    name_ As String
    input As String
End Type

Public Sub GenerateExamples()
    Dim e As Examples
    e = NewExamples()
    AddExample e, NewExample("pretty-function", _
        "=CONCAT(""R"",MOD(ROW()-6,2)*2+1,""C"",INT((ROW()-6)/2)*2+1)" _
    )
    Save e, Left(ThisWorkbook.Path, InStrRev(ThisWorkbook.Path, "\bin")) & "examples"
End Sub

Private Function NewExamples() As Examples
    Dim e As Examples
    Set e.val = New Collection
    NewExamples = e
End Function

Private Function NewExample(name_ As String, input_ As String) As Example
    Dim e As Example
    e.name_ = name_
    e.input = input_
    NewExample = e
End Function

Private Sub AddExample(e As Examples, example_ As Example)
    e.val.Add Array(example_.name_, example_.input)
End Sub

Private Sub Save(e As Examples, dirPath As String)
    ' MkDir will error if the folder already exists,
    On Error Resume Next
        MkDir dirPath
    On Error GoTo 0

    Dim example_ As Variant
    For Each example_ In e.val
        Dim fileNumber As Long
        Dim filePath As String
        fileNumber = FreeFile
        filePath = dirPath & "\" & example_(0) & ".txt"
        Open filePath For Output As #fileNumber
            Dim content As String
            content = Formulas.Pretty(CStr(example_(1)), _
                Formulas.NewFormatter( _
                    indent:=" ", _
                    indentLength:=2, _
                    newLine:=vbCrLf, _
                    eqAtStart:=True, _
                    newLineAtEof:=True _
                ) _
            )
            Print #fileNumber, "in:"
            Print #fileNumber, CStr(example_(1))
            Print #fileNumber, ' Print blank line to file
            Print #fileNumber, "out:"
            Print #fileNumber, content
        Close #fileNumber
    Next example_
End Sub
