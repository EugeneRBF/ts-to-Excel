VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Lexem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' (C) 2021 Eugene Torkhov
'
Option Explicit

Private tsString As String, tsIndex As Long, startTerm As Long
Private Keywords As New Collection, separators As New Dictionary, anyStop As String, anySpace As String

Private typ As LexemType
Private val As String

Public Enum LexemType
    Keyword
    Identifier
    StringValue
    Comment
    Comma
    SemiColon
    Colon
    Assign
    SquareStart
    SquareEnd
    CurlyStart
    CurlyEnd
    EOF
    kwConst
    kwExport
    kwImport
End Enum

Public Property Get value() As String
    value = val
End Property

Public Property Get LexemType() As LexemType
    LexemType = typ
End Property

Public Property Get index() As Long
    index = startTerm
End Property

Public Property Get toString() As String
    toString = toName & ", " & val
End Property

Public Function toName()
    Select Case typ
        Case Keyword:
            toName = "Keyword"
        Case Identifier:
            toName = "Identifier"
        Case StringValue:
            toName = "StringValue"
        Case Comment:
            toName = "Comment"
        Case Comma:
            toName = "Comma"
        Case Colon:
            toName = "Colon"
        Case SemiColon:
            toName = "SemiColon"
        Case Colon:
            toName = "Colon"
        Case Assign:
            toName = "Assign"
        Case SquareStart:
            toName = "SquareStart"
        Case SquareEnd:
            toName = "SquareEnd"
        Case CurlyStart:
            toName = "CurlyStart"
        Case CurlyEnd:
            toName = "CurlyEnd"
        Case EOF:
            toName = "EOF"
        Case kwConst
            toName = "KW Const"
        Case kwExport:
            toName = "KW Export"
        Case kwImport:
            toName = "KW Import"
        Case Identifier:
            toName = "Identifier"
        Case Else:
            toName = "--undefined--"
    End Select
End Function

Public Sub init(source As String)
    Dim ix As Integer
    
    tsString = source
    tsIndex = 1
    
    Keywords.add "const"
    Keywords.add "import"
    
    separators.add 1, " "
    separators.add 2, vbTab
    separators.add 3, ","
    separators.add 4, ":"
    separators.add 5, ";"
    separators.add 6, "="
    separators.add 7, "["
    separators.add 8, "]"
    separators.add 9, "{"
    separators.add 10, "}"
    separators.add 11, "'"
    separators.add 12, """"
    separators.add 13, "`"
    
    anyStop = ""
    For ix = 1 To separators.Count
        anyStop = anyStop & separators.Item(ix)
    Next ix

    anySpace = " " & vbTab & vbNewLine
End Sub

Private Sub Class_Terminate()
    Set Keywords = New Collection
    separators.RemoveAll
End Sub

Public Function parse() As Lexem
    Dim ix As Integer, k As Variant
    getToken
    Set parse = Me
    If typ = Keyword Then
        Select Case val
        Case "const":
            typ = kwConst
        'Case "export":
        '    typ = kwExport
        Case "import":
            typ = kwImport
        Case Else
            typ = Identifier
        End Select
    End If
End Function

Private Function skipSpace()
    Dim ix As Integer, catched As Boolean
    ' Increment index to skip over spaces
    Do While tsIndex > 0 And tsIndex <= VBA.Len(tsString)
        catched = VBA.InStr(1, anySpace, VBA.Mid$(tsString, tsIndex, 1)) > 0
        If catched Then
            Exit Do
        End If
        tsIndex = tsIndex + 1
    Loop
End Function


Private Function getToken() As Lexem
    Dim ix As Integer, catched As Boolean
    
    typ = EOF
    val = ""
    Do While tsIndex > 0 And tsIndex < VBA.Len(tsString)
        catched = VBA.InStr(1, anySpace, VBA.Mid$(tsString, tsIndex, 1)) > 0
        If Not catched Then
            Exit Do
        End If
        tsIndex = tsIndex + 1
    Loop
    If catched Then
        typ = EOF
        GoTo BuildResult
    End If
    startTerm = tsIndex
    Do While tsIndex > 0 And tsIndex < VBA.Len(tsString)
        catched = VBA.InStr(1, anyStop, VBA.Mid$(tsString, tsIndex, 1)) > 0
        If catched Then
            Exit Do
        End If
        tsIndex = tsIndex + 1
    Loop
    If tsIndex - startTerm = 0 Then
        val = VBA.Mid$(tsString, startTerm, 1)
        If VBA.InStr(1, """'`", val) > 0 And tsIndex < VBA.Len(tsString) Then
            tsIndex = tsIndex + 1
            val = findEndOfQuote(val)
            'val = VBA.Mid$(tsString, startTerm + 1, tsIndex - 1 - startTerm)
            typ = StringValue
        Else
            Select Case val
            Case ";"
                typ = SemiColon
            Case "="
                typ = Assign
            Case "["
                typ = SquareStart
            Case "]"
                typ = SquareEnd
            Case "{"
                typ = CurlyStart
            Case "}"
                typ = CurlyEnd
            Case ","
                typ = Comma
            Case ":"
                typ = Colon
            Case Else
                typ = EOF
            End Select
            val = VBA.Mid$(tsString, startTerm, 1)
        End If
        tsIndex = tsIndex + 1
        GoTo BuildResult
    ElseIf VBA.InStr(VBA.Mid$(tsString, startTerm, tsIndex - startTerm), "/*") = 1 Then
        typ = Comment
        skipToCommentClose
        val = VBA.Mid$(tsString, startTerm, tsIndex - startTerm + 1)
        tsIndex = tsIndex + 1
    ElseIf VBA.InStr(VBA.Mid$(tsString, startTerm, tsIndex - startTerm), "//") = 1 Then
        typ = Comment
        skipToEndOfLine
        val = VBA.Mid$(tsString, startTerm, tsIndex - startTerm + 1)
        tsIndex = tsIndex + 1
    Else
        typ = Keyword
        val = VBA.Mid$(tsString, startTerm, tsIndex - startTerm)
    End If

BuildResult:
    Set getToken = Me
End Function

Private Function findEndOfQuote(ByVal quote As String) As String
    Dim catched As Integer, startPart As Long
    startPart = tsIndex
    findEndOfQuote = ""
    Do While tsIndex > 0 And tsIndex <= VBA.Len(tsString)
        catched = VBA.InStr(1, quote & vbNewLine & "\", VBA.Mid$(tsString, tsIndex, 1))
        If catched > 0 Then
            If catched >= 3 Then
                findEndOfQuote = findEndOfQuote & VBA.Mid$(tsString, startPart, tsIndex - startPart)
                tsIndex = tsIndex + 1
                startPart = tsIndex
            ElseIf catched = 2 Then
                err.Raise Number:=vbObjectError + 555, Description:="¬стретилс€ перевод строки внутри строковой константы"
            Else
                findEndOfQuote = findEndOfQuote & VBA.Mid$(tsString, startPart, tsIndex - startPart)
                Exit Function
            End If
        End If
        tsIndex = tsIndex + 1
    Loop
    err.Raise Number:=vbObjectError + 556, Description:="Ќе найдено завершение строковой константы"
End Function

Private Sub skipToEndOfLine()
    Dim catched As Long
    catched = VBA.InStr(tsIndex, tsString, vbNewLine)
    If catched > 0 Then
        tsIndex = catched - 1
    Else
        tsIndex = VBA.Len(tsString)
    End If
End Sub

Private Sub skipToCommentClose()
    Dim catched As Long
    catched = VBA.InStr(tsIndex, tsString, "*/")
    If catched > 0 Then
        tsIndex = catched + 1
    Else
        tsIndex = VBA.Len(tsString)
    End If
End Sub

