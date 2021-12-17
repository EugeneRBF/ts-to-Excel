'
' (C) 2021 Eugene Torkhov
'
Option Explicit On

Public Enum Action
    None
    ReduceString
    ReduceObject
    Copy
    Copy2
    Merge
    GetLocale
    Finish
    FinishAll
End Enum

Public Class SyntaxRule

    Private tokens As New Collection
    Private act As Action
    Private tgt As TokenType
    Private val As New Collection

    Public Function add(ByVal t As TokenType) As SyntaxRule
        tokens.Add(t)
        add = Me
    End Function

    Property canDo() As Action
        Get
            Return act
        End Get
        Set(ByVal value As Action)
            act = value
        End Set
    End Property

    Property target() As TokenType
        Get
            Return tgt
        End Get
        Set(ByVal value As TokenType)
            tgt = value
        End Set
    End Property

    Property value() As Collection
        Get
            Return val
        End Get
        Set(ByVal value As Collection)
            val = value
        End Set
    End Property

    ReadOnly Property items() As Collection
        Get
            If tokens Is Nothing Then tokens = New Collection
            Return tokens
        End Get
    End Property

    Public Function toStr() As String
        Dim ix As Integer
        toStr = "target: " & tokenName(tgt)
        toStr = toStr & ", tokens:"
        For ix = 1 To tokens.Count
            toStr = toStr & " " & tokenName(tokens.Item(ix)) & ","
        Next ix
        If act = Action.None Then
            toStr = toStr & " action: skip"
        Else
            toStr = toStr & " action: reduce"
        End If
    End Function

    Private Function tokenName(ByVal tgt As TokenType) As String
        Select Case tgt
            Case TokenType.Programm
                tokenName = "Programm"
            Case TokenType.Statement
                tokenName = "Statement"
            Case TokenType.StatementList
                tokenName = "StatementList"
            Case TokenType.oProperty
                tokenName = "Property"
            Case TokenType.oPropertyList
                tokenName = "PropertyList"
            Case TokenType.tImport
                tokenName = "import"
            Case TokenType.tIdentifier
                tokenName = "Identifier"
            Case TokenType.tIdentList
                tokenName = "IdentList"
            Case TokenType.tAssign
                tokenName = ":="
            Case TokenType.tColon
                tokenName = "Colon"
            Case TokenType.tComma
                tokenName = "Comma"
            Case TokenType.tSemiColon
                tokenName = ";"
            Case TokenType.tCurlyStart
                tokenName = "{"
            Case TokenType.tCurlyEnd
                tokenName = "}"
            Case TokenType.tBrackedStart
                tokenName = "["
            Case TokenType.tBrackedEnd
                tokenName = "]"
            Case TokenType.tComment
                tokenName = "Comment"
            Case TokenType.vValueList
                tokenName = "ValueList"
            Case TokenType.vObject
                tokenName = "Object"
            Case TokenType.vArray
                tokenName = "Array"
            Case TokenType.vString
                tokenName = "String"
            Case TokenType.EOP
                tokenName = "EOF"
            Case Else
                tokenName = "-unknown-"
        End Select

    End Function


    Public Function create(ByVal canDo_ As Action, ByVal t As TokenType) As SyntaxRule
        Dim rule As New SyntaxRule
        rule.canDo = canDo_
        rule.target = t
        create = rule
    End Function
End Class
