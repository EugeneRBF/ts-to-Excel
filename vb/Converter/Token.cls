'
' (C) 2021 Eugene Torkhov
'
Option Explicit

Public Enum TokenType
    Programm
    Statement
    StatementList
    oProperty
    oPropertyList
    tImport
    tIdentifier
    tIdentList
    tAssign
    tColon
    tComma
    tSemiColon
    tCurlyStart
    tCurlyEnd
    tBrackedStart
    tBrackedEnd
    tComment
    vValueList
    vObject
    vArray
    vString
    EOP
End Enum

Public Class Token

    Private typ As TokenType
    Private items As New Collection
    Private loc As String

    Property type() As TokenType
        Get
            Return typ
        End Get
        Set(ByVal value As TokenType)
            typ = value
        End Set
    End Property

    Property value() As Collection
        Get
            Return items
        End Get
        Set(ByVal value As Collection)
            items = value
        End Set
    End Property

    Property locale() As String
        Get
            Return loc
        End Get
        Set(ByVal value As String)
            loc = value
        End Set
    End Property


    Public Function toStr() As String
        Select Case typ
            Case TokenType.Programm
                toStr = "Programm"
            Case TokenType.Statement
                toStr = "Statement"
            Case TokenType.StatementList
                toStr = "StatementList"
            Case TokenType.oProperty
                toStr = "Property"
            Case TokenType.oPropertyList
                toStr = "PropertyList"
            Case TokenType.tImport
                toStr = "import"
            Case TokenType.tIdentifier
                toStr = "Identifier"
            Case TokenType.tAssign
                toStr = ":="
            Case TokenType.tColon
                toStr = ":"
            Case TokenType.tComma
                toStr = ","
            Case TokenType.tSemiColon
                toStr = ":"
            Case TokenType.tCurlyStart
                toStr = "{"
            Case TokenType.tCurlyEnd
                toStr = "}"
            Case TokenType.tBrackedStart
                toStr = "["
            Case TokenType.tBrackedEnd
                toStr = "]"
            Case TokenType.tComment
                toStr = "Comment"
            Case TokenType.vValueList
                toStr = "ValueList"
            Case TokenType.vObject
                toStr = "Object"
            Case TokenType.vArray
                toStr = "Array"
            Case TokenType.vString
                toStr = "String"
            Case TokenType.EOP
                toStr = "EOF"
        End Select
        Return ToString()
    End Function

    Public Function toToken(ByVal lex As Lexem) As Token
        Dim tok As New Token

        Select Case lex.lexemType
            Case LexemType.Keyword
                Err.Raise(Number:=vbObjectError + 1000, Description:="Не известное ключевое слово: " & lex.value)
            Case LexemType.Identifier
                tok.type = TokenType.tIdentifier
            Case LexemType.StringValue
                tok.type = TokenType.vString
            Case LexemType.Comment
                tok.type = TokenType.tComment
            Case LexemType.Comma
                tok.type = TokenType.tComma
            Case LexemType.SemiColon
                tok.type = TokenType.tSemiColon
            Case LexemType.Colon
                tok.type = TokenType.tColon
            Case LexemType.Assign
                tok.type = TokenType.tAssign
            Case LexemType.SquareStart
                tok.type = TokenType.tBrackedStart
            Case LexemType.SquareEnd
                tok.type = TokenType.tBrackedEnd
            Case LexemType.CurlyStart
                tok.type = TokenType.tCurlyStart
            Case LexemType.CurlyEnd
                tok.type = TokenType.tCurlyEnd
            Case LexemType.EOF
                tok.type = TokenType.EOP
                'Case kwConst
                '    tok.setType = tConst
                'Case kwExport
                '    tok.setType = tExport
            Case LexemType.kwImport
                tok.type = TokenType.tImport
        End Select
        tok.value.Add(lex.value)
        toToken = tok
    End Function
End Class
