VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SyntaxRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' (C) 2021 Eugene Torkhov
'
Private tokens As New Collection
Private act As Action
Private tgt As TokenType
Private val As New Collection

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

Public Function add(t As TokenType) As SyntaxRule
    tokens.add t
    Set add = Me
End Function

Property Get canDo() As Action
    canDo = act
End Property

Property Let canDo(what As Action)
    act = what
End Property

Property Get target() As TokenType
    target = tgt
End Property

Property Let target(what As TokenType)
    tgt = what
End Property

Property Get value() As Collection
    Set value = val
End Property

Property Set value(what As Collection)
    Set val = what
End Property

Property Get items() As Collection
    If tokens Is Nothing Then Set tokens = New Collection
    Set items = tokens
End Property

Public Function toString() As String
    Dim ix As Integer
    toString = "target: " & tokenName(tgt)
    toString = toString & ", tokens:"
    For ix = 1 To tokens.Count
        toString = toString & " " & tokenName(tokens.Item(ix)) & ","
    Next ix
    If act = None Then
        toString = toString & " action: skip"
    Else
        toString = toString & " action: reduce"
    End If
End Function

Private Function tokenName(tgt As TokenType) As String
    Select Case tgt
    Case Programm
        tokenName = "Programm"
    Case Statement
        tokenName = "Statement"
    Case StatementList
        tokenName = "StatementList"
    Case oProperty
        tokenName = "Property"
    Case oPropertyList
        tokenName = "PropertyList"
    Case tImport
        tokenName = "import"
    Case tExport
        tokenName = "export"
    Case tConst
        tokenName = "const"
    Case tIdentifier
        tokenName = "Identifier"
    Case tAssign
        tokenName = ":="
    Case tColon
        tokenName = "Colon"
    Case tComma
        tokenName = "Comma"
    Case tSemiColon
        tokenName = ";"
    Case tCurlyStart
        tokenName = "{"
    Case tCurlyEnd
        tokenName = "}"
    Case tBrackedStart
        tokenName = "["
    Case tBrackedEnd
        tokenName = "]"
    Case tComment
        tokenName = "Comment"
    Case vValueList
        tokenName = "ValueList"
    Case vObject
        tokenName = "Object"
    Case vArray
        tokenName = "Array"
    Case vString
        tokenName = "String"
    Case EOP
        tokenName = "EOF"
    Case exError 'для генерации ошибок разбора
        tokenName = "Error"
    End Select
    
End Function


Public Function create(canDo_ As Action, t As TokenType) As SyntaxRule
    Dim rule As New SyntaxRule
    rule.canDo = canDo_
    rule.target = t
    Set create = rule
End Function
