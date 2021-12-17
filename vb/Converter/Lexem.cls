'
' (C) 2021 Eugene Torkhov
'
Option Explicit On

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
    kwImport
End Enum

Public Class Lexem

    Private tsString As String, tsIndex As Long, startTerm As Long
    Private Keywords As New Collection, separators As New Dictionary(Of String, String), anyStop As String, anySpace As String

    Private typ As LexemType
    Private val As String

    Public ReadOnly Property value() As String
        Get
            Return val
        End Get
    End Property

    Public ReadOnly Property lexemType() As LexemType
        Get
            Return typ
        End Get
    End Property

    Public ReadOnly Property index() As Long
        Get
            Return tsIndex
        End Get
    End Property

    Public ReadOnly Property toStr() As String
        Get
            Return toName() & ", " & val
        End Get

    End Property

    Public Function toName()
        Select Case typ
            Case LexemType.Keyword
                toName = "Keyword"
            Case LexemType.Identifier
                toName = "Identifier"
            Case LexemType.StringValue
                toName = "StringValue"
            Case LexemType.Comment
                toName = "Comment"
            Case LexemType.Comma
                toName = "Comma"
            Case LexemType.Colon
                toName = "Colon"
            Case LexemType.SemiColon
                toName = "SemiColon"
            Case LexemType.Colon
                toName = "Colon"
            Case LexemType.Assign
                toName = "Assign"
            Case LexemType.SquareStart
                toName = "SquareStart"
            Case LexemType.SquareEnd
                toName = "SquareEnd"
            Case LexemType.CurlyStart
                toName = "CurlyStart"
            Case LexemType.CurlyEnd
                toName = "CurlyEnd"
            Case LexemType.EOF
                toName = "EOF"
            Case LexemType.kwImport
                toName = "KW Import"
            Case LexemType.Identifier
                toName = "Identifier"
            Case Else
                toName = "--undefined--"
        End Select
    End Function

    Public Sub init(ByRef source As String)
        Dim ix As Integer

        tsString = source
        tsIndex = 1

        Keywords.Add("import")

        separators.add(1, " ")
        separators.add(2, vbTab)
        separators.add(3, ",")
        separators.add(4, ":")
        separators.add(5, ";")
        separators.add(6, "=")
        separators.add(7, "[")
        separators.add(8, "]")
        separators.add(9, "{")
        separators.add(10, "}")
        separators.add(11, "'")
        separators.add(12, """")
        separators.add(13, "`")

        anyStop = ""
        For ix = 1 To separators.Count
            anyStop = anyStop & separators.Item(ix)
        Next ix

        anySpace = " " & vbTab & vbNewLine
    End Sub

    Protected Overrides Sub Finalize()
        Keywords = New Collection
        separators.Clear()
    End Sub

    Public Function parse() As Lexem
        getToken()
        parse = Me
        If typ = lexemType.Keyword Then
            Select Case val
                'Case "const":
                '    typ = kwConst
                'Case "export":
                '    typ = kwExport
                Case "import"
                    typ = lexemType.kwImport
                Case Else
                    typ = lexemType.Identifier
            End Select
        End If
    End Function

    Private Sub skipSpace()
        Dim catched As Boolean
        ' Increment index to skip over spaces
        Do While tsIndex > 0 And tsIndex <= Len(tsString)
            catched = InStr(1, anySpace, Mid$(tsString, tsIndex, 1)) > 0
            If catched Then
                Exit Do
            End If
            tsIndex = tsIndex + 1
        Loop
    End Sub


    Private Function getToken() As Lexem
        Dim catched As Boolean

        typ = lexemType.EOF
        val = ""
        Do While tsIndex > 0 And tsIndex < Len(tsString)
            catched = InStr(1, anySpace, Mid$(tsString, tsIndex, 1)) > 0
            If Not catched Then
                Exit Do
            End If
            tsIndex = tsIndex + 1
        Loop
        If catched Then
            typ = lexemType.EOF
            GoTo BuildResult
        End If
        startTerm = tsIndex
        Do While tsIndex > 0 And tsIndex < Len(tsString)
            catched = InStr(1, anyStop, Mid$(tsString, tsIndex, 1)) > 0
            If catched Then
                Exit Do
            End If
            tsIndex = tsIndex + 1
        Loop
        If tsIndex - startTerm = 0 Then
            val = Mid$(tsString, startTerm, 1)
            If InStr(1, """'`", val) > 0 And tsIndex < Len(tsString) Then
                tsIndex = tsIndex + 1
                val = findEndOfQuote(val)
                'val = Mid$(tsString, startTerm + 1, tsIndex - 1 - startTerm)
                typ = lexemType.StringValue
            Else
                Select Case val
                    Case ";"
                        typ = lexemType.SemiColon
                    Case "="
                        typ = lexemType.Assign
                    Case "["
                        typ = lexemType.SquareStart
                    Case "]"
                        typ = lexemType.SquareEnd
                    Case "{"
                        typ = lexemType.CurlyStart
                    Case "}"
                        typ = lexemType.CurlyEnd
                    Case ","
                        typ = lexemType.Comma
                    Case ":"
                        typ = lexemType.Colon
                    Case Else
                        typ = lexemType.EOF
                End Select
                val = Mid$(tsString, startTerm, 1)
            End If
            tsIndex = tsIndex + 1
            GoTo BuildResult
        ElseIf InStr(Mid$(tsString, startTerm, tsIndex - startTerm), "/*") = 1 Then
            typ = lexemType.Comment
            skipToCommentClose()
            val = Mid$(tsString, startTerm, tsIndex - startTerm + 1)
            tsIndex = tsIndex + 1
        ElseIf InStr(Mid$(tsString, startTerm, tsIndex - startTerm), "//") = 1 Then
            typ = lexemType.Comment
            skipToEndOfLine()
            val = Mid$(tsString, startTerm, tsIndex - startTerm + 1)
            tsIndex = tsIndex + 1
        Else
            typ = lexemType.Keyword
            val = Mid$(tsString, startTerm, tsIndex - startTerm)
        End If

BuildResult:
        getToken = Me
    End Function

    Private Function findEndOfQuote(ByVal quote As String) As String
        Dim catched As Integer, startPart As Long
        startPart = tsIndex
        findEndOfQuote = ""
        Do While tsIndex > 0 And tsIndex <= Len(tsString)
            catched = InStr(1, quote & vbNewLine & "\", Mid$(tsString, tsIndex, 1))
            If catched > 0 Then
                If catched >= 3 Then
                    findEndOfQuote = findEndOfQuote & Mid$(tsString, startPart, tsIndex - startPart)
                    tsIndex = tsIndex + 1
                    startPart = tsIndex
                ElseIf catched = 2 Then
                    Err.Raise(Number:=vbObjectError + 555, Description:="¬стретилс€ перевод строки внутри строковой константы")
                Else
                    findEndOfQuote = findEndOfQuote & Mid$(tsString, startPart, tsIndex - startPart)
                    Exit Function
                End If
            End If
            tsIndex = tsIndex + 1
        Loop
        Err.Raise(Number:=vbObjectError + 556, Description:="Ќе найдено завершение строковой константы")
    End Function

    Private Sub skipToEndOfLine()
        Dim catched As Long
        catched = InStr(tsIndex, tsString, vbNewLine, CompareMethod.Binary)
        If catched > 0 Then
            tsIndex = catched - 1
        Else
            tsIndex = Len(tsString)
        End If
    End Sub

    Private Sub skipToCommentClose()
        Dim catched As Long
        catched = InStr(tsIndex, tsString, "*/", CompareMethod.Binary)
        If catched > 0 Then
            tsIndex = catched + 1
        Else
            tsIndex = Len(tsString)
        End If
    End Sub

End Class