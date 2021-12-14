VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Token"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' (C) 2021 Eugene Torkhov
'
Option Explicit

Private typ As TokenType
Private items As New Collection
Private loc As String

Public Enum TokenType
    Programm
    Statement
    StatementList
    oProperty
    oPropertyList
    tImport
    tExport
    tConst
    tIdentifier
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

Property Let setType(t As TokenType)
    typ = t
End Property

Property Get getType() As TokenType
    getType = typ
End Property

Property Let value(obj As Variant)
    items.add obj
End Property

Property Get value() As Collection
    Set value = items
End Property

Property Get locale() As String
    locale = loc
End Property

Property Let locale(loc_ As String)
    loc = loc_
End Property

Public Function toString() As String
    Select Case typ
    Case Programm
        toString = "Programm"
    Case Statement
        toString = "Statement"
    Case StatementList
        toString = "StatementList"
    Case oProperty
        toString = "Property"
    Case oPropertyList
        toString = "PropertyList"
    Case tImport
        toString = "import"
    Case tExport
        toString = "export"
    Case tConst
        toString = "const"
    Case tIdentifier
        toString = "Identifier"
    Case tAssign
        toString = ":="
    Case tColon
        toString = ":"
    Case tComma
        toString = ","
    Case tSemiColon
        toString = ":"
    Case tCurlyStart
        toString = "{"
    Case tCurlyEnd
        toString = "}"
    Case tBrackedStart
        toString = "["
    Case tBrackedEnd
        toString = "]"
    Case tComment
        toString = "Comment"
    Case vValueList
        toString = "ValueList"
    Case vObject
        toString = "Object"
    Case vArray
        toString = "Array"
    Case vString
        toString = "String"
    Case EOP
        toString = "EOF"
    End Select
    
End Function

Public Static Function toToken(lex As Lexem) As Token
    Dim tok As New Token
    
    Select Case lex.LexemType
    Case Keyword
        err.Raise Number:=vbObjectError + 1000, Description:="Не известное ключевое слово: " & lex.value
    Case Identifier
        tok.setType = tIdentifier
    Case StringValue
        tok.setType = vString
    Case Comment
        tok.setType = tComment
    Case Comma
        tok.setType = tComma
    Case SemiColon
        tok.setType = tSemiColon
    Case Colon
        tok.setType = tColon
    Case Assign
        tok.setType = tAssign
    Case SquareStart
        tok.setType = tBrackedStart
    Case SquareEnd
        tok.setType = tBrackedEnd
    Case CurlyStart
        tok.setType = tCurlyStart
    Case CurlyEnd
        tok.setType = tCurlyEnd
    Case EOF
        tok.setType = EOP
    Case kwConst
        tok.setType = tConst
    Case kwExport
        tok.setType = tExport
    Case kwImport
        tok.setType = tImport
    End Select
    tok.value = lex.value
    Set toToken = tok
End Function
