Attribute VB_Name = "tsParse"
'
' (C) 2021 Eugene Torkhov
'
Option Explicit

Dim rules As Collection
Dim currentCol As Long, maxKeys As Long

Public Sub Main()
    'On Error Resume Next
    Menu_ParsingTS_OnAction
End Sub


Public Sub Menu_ParsingTS_OnAction()
    Dim myFile As String, stream As New ADODB.stream, ix As Long, text As String
    Dim parser As New Lexem, tkn As New Token, stack As New Collection
    
    prepareRules
    currentCol = 1

    With Application.FileDialog(msoFileDialogOpen)
        .AllowMultiSelect = True
        .Show
 
        For ix = 1 To .SelectedItems.Count
            ' Load each selected file
            stream.Charset = "utf-8"
            stream.Open
            stream.LoadFromFile .SelectedItems(ix)
            text = text & stream.ReadText() & vbNewLine
            stream.Close
        Next ix
     End With
    
    If VBA.Len(text) > 0 Then
        
        parser.init text
        
        Do
            parser.parse
            'Debug.Print parser.index, parser.toString()
            Set tkn = tkn.toToken(parser)
            stack.add Item:=tkn
            tryReduce stack
        Loop While parser.LexemType <> EOF 'And jsonIndex < VBA.Len(text)
        
        For ix = stack.Count To 1 Step -1
            stack.Remove ix
        Next ix
        MsgBox "Complete"
    
    End If
    

End Sub

Private Sub prepareRules()
    Dim rule As New SyntaxRule
    
    Set rules = New Collection
    
    rules.add Item:=rule.create(FinishAll, Programm).add(StatementList).add(EOP)
    rules.add Item:=rule.create(Copy, Statement).add(Statement).add(tSemiColon)             ' stmt ;
    rules.add Item:=rule.create(None, Statement).add(tIdentifier).add(tAssign).add(vString) ' ident = 'string'
    rules.add Item:=rule.create(None, Statement).add(tImport).add(vString)                  ' import 'string'
    rules.add Item:=rule.create(None, Statement).add(tImport).add(tCurlyStart).add(tIdentifier).add(tCurlyEnd).add(tIdentifier).add(vString) ' import { class } from 'string'
    rules.add Item:=rule.create(None, Statement).add(tImport).add(tCurlyStart).add(tIdentList).add(tCurlyEnd).add(tIdentifier).add(vString) ' import { class } from 'string'
    rules.add Item:=rule.create(None, Statement).add(tImport).add(tIdentifier).add(vString) ' import * from 'string'
    rules.add Item:=rule.create(None, tIdentList).add(tIdentifier).add(tComma).add(tIdentifier)     ' identList := <ident>, <ident> ;
    rules.add Item:=rule.create(None, tIdentList).add(tIdentList).add(tComma).add(tIdentifier)     ' identList := <ident>, <ident> ;
    rules.add Item:=rule.create(Merge, StatementList).add(StatementList).add(Statement)     ' stmtList := <stmtList> <stmt>
    rules.add Item:=rule.create(Copy, StatementList).add(StatementList).add(tSemiColon)     ' stmtList := <stmtList> ;
    rules.add Item:=rule.create(Copy, StatementList).add(Statement)                         ' stmtList := <stmt>
    rules.add Item:=rule.create(Finish, Statement).add(tIdentifier).add(tAssign).add(vObject)   ' stmt := <ident> = <object>
    rules.add Item:=rule.create(GetLocale, tIdentifier).add(tIdentifier).add(tBrackedStart).add(vString).add(tBrackedEnd)   ' ident = <ident> [ 'string' ]
    rules.add Item:=rule.create(Copy2, tIdentifier).add(tIdentifier).add(tIdentifier)        ' ident := export const <ident>
    rules.add Item:=rule.create(Copy2, vObject).add(tCurlyStart).add(oPropertyList).add(tCurlyEnd)
    rules.add Item:=rule.create(None, vObject).add(tCurlyStart).add(tCurlyEnd)
    rules.add Item:=rule.create(ReduceString, oProperty).add(tIdentifier).add(tColon).add(vString)
    rules.add Item:=rule.create(ReduceObject, oProperty).add(tIdentifier).add(tColon).add(vObject)
    rules.add Item:=rule.create(ReduceString, oProperty).add(vString).add(tColon).add(vString)
    rules.add Item:=rule.create(ReduceObject, oProperty).add(vString).add(tColon).add(vObject)
    rules.add Item:=rule.create(Copy, oPropertyList).add(oPropertyList).add(tComma)
    rules.add Item:=rule.create(Merge, oPropertyList).add(oPropertyList).add(oProperty)
    rules.add Item:=rule.create(Copy, oPropertyList).add(oProperty)
    rules.add Item:=rule.create(Copy, vString).add(vString).add(tComment)
    rules.add Item:=rule.create(None, tComma).add(tComma).add(tComment)
    rules.add Item:=rule.create(Copy, vObject).add(vObject).add(tComment)
    rules.add Item:=rule.create(None, tAssign).add(tAssign).add(tComment)
    rules.add Item:=rule.create(None, tBrackedStart).add(tBrackedStart).add(tComment)
    rules.add Item:=rule.create(None, tBrackedEnd).add(tBrackedEnd).add(tComment)
    rules.add Item:=rule.create(None, tCurlyStart).add(tCurlyStart).add(tComment)
    rules.add Item:=rule.create(None, tCurlyEnd).add(tCurlyEnd).add(tComment)
    rules.add Item:=rule.create(Copy, tIdentifier).add(tIdentifier).add(tComment)
    rules.add Item:=rule.create(Copy, Statement).add(Statement).add(tComment)
    rules.add Item:=rule.create(Copy, StatementList).add(StatementList).add(tComment)
    'rules.add item:=rule.create(Copy2, StatementList).add(tComment).add(StatementList)
    rules.add Item:=rule.create(None, Programm).add(Programm).add(tComment)
    rules.add Item:=rule.create(Copy, oProperty).add(oProperty).add(tComment)
    rules.add Item:=rule.create(Copy, oPropertyList).add(oPropertyList).add(tComment)
End Sub

Private Sub tryReduce(ByRef stack As Collection)
    Dim ix As Long, rule As SyntaxRule, ss As Long, iy As Integer, matched As Boolean, newTok As Token, t As Token
    Dim key As String, val As String, kv As KeyValPair, col As Collection, tr As Translation, wasReduce As Boolean
    'On Error Resume Next
    Do
        wasReduce = False
        For ix = 1 To rules.Count
            Set rule = rules.Item(ix)
            If stack.Count >= rule.items.Count Then
                ss = stack.Count - rule.items.Count
                matched = True
                For iy = 1 To rule.items.Count
                    Set t = stack.Item(ss + iy)
                    If t.getType <> rule.items.Item(iy) Then
                        matched = False
                        Exit For
                    End If
                Next iy
                If matched Then
                    Set col = New Collection
                    Set kv = New KeyValPair
                    Set tr = New Translation
                    Select Case rule.canDo
                    Case ReduceString: beforeReduceString stack, rule, kv
                    Case ReduceObject: beforeReduceObject stack, rule, col
                    Case Copy: beforeCopy stack, rule, col
                    Case Copy2: beforeCopy2 stack, rule, col
                    Case Merge: beforeMerge stack, rule, col
                    Case GetLocale: beforeGetLocale stack, rule, tr
                    Case Finish: beforeFinish stack, rule, tr
                    Case FinishAll: beforeFinishAll stack, rule, col
                    End Select
                    
                    'Debug.Print "Stack reduced to " & rule.toString
                    For iy = stack.Count To ss + 1 Step -1
                        stack.Remove iy
                    Next iy
                    
                    Set newTok = New Token
                    newTok.setType = rule.target
                    stack.add newTok
                    
                    wasReduce = True
                    
                    Select Case rule.canDo
                    Case ReduceString: afterReduceString newTok, kv
                    Case ReduceObject: afterReduceObject newTok, col
                    Case Copy: afterCopy newTok, col
                    Case Copy2: afterCopy newTok, col
                    Case Merge: afterMerge newTok, col
                    Case GetLocale: afterGetLocale newTok, tr
                    Case Finish: afterFinish newTok, tr
                    Case FinishAll: afterFinishAll newTok, col
                    End Select
                    
                    Set col = Nothing
                    Set kv = Nothing
                    Set tr = Nothing
                    'tryReduce stack        'in recursive mode or just leave from the loop
                    Exit For
                End If
            End If
        Next ix
    Loop While wasReduce
    On Error GoTo 0
End Sub

Private Sub beforeReduceString(ByRef stack As Collection, ByRef rule As SyntaxRule, ByRef kv As KeyValPair)
    Dim ss As Long
    ss = stack.Count - rule.items.Count
    kv.key = stack.Item(ss + 1).value.Item(1)
    kv.val = stack.Item(ss + 3).value.Item(1)
End Sub

Private Sub afterReduceString(ByRef newTok As Token, ByRef kv As KeyValPair)
    newTok.value = kv
End Sub

Private Sub beforeCopy(ByRef stack As Collection, ByRef rule As SyntaxRule, ByRef col As Collection)
    Dim ss As Long, v As Object, ix As Long
    ss = stack.Count - rule.items.Count
    Set v = stack.Item(ss + 1).value
    For ix = 1 To v.Count
        col.add v.Item(ix)
    Next ix
End Sub

Private Sub afterCopy(ByRef newTok As Token, ByRef col As Collection)
    Dim ix As Long
    For ix = 1 To col.Count
        newTok.value = col.Item(ix)
    Next ix
End Sub

Private Sub beforeCopy2(ByRef stack As Collection, ByRef rule As SyntaxRule, ByRef col As Collection)
    Dim ss As Long, v As Object, ix As Long
    ss = stack.Count - rule.items.Count
    Set v = stack.Item(ss + 2).value
    For ix = 1 To v.Count
        col.add v.Item(ix)
    Next ix
End Sub

Private Sub beforeMerge(ByRef stack As Collection, ByRef rule As SyntaxRule, ByRef col As Collection)
    Dim ss As Long, v As Object, ix As Long
    ss = stack.Count - rule.items.Count
    Set v = stack.Item(ss + 1).value
    For ix = 1 To v.Count
        col.add v.Item(ix)
    Next ix
    Set v = stack.Item(ss + 2).value
    For ix = 1 To v.Count
        col.add v.Item(ix)
    Next ix
End Sub

Private Sub afterMerge(ByRef newTok As Token, ByRef col As Collection)
    Dim ix As Long
    For ix = 1 To col.Count
        newTok.value = col.Item(ix)
    Next ix
End Sub


Private Sub beforeReduceObject(ByRef stack As Collection, ByRef rule As SyntaxRule, ByRef col As Collection)
    Dim ss As Long, v As Object, ix As Long, key As String, kv As KeyValPair
    ss = stack.Count - rule.items.Count
    Set v = stack.Item(ss + 1).value
    key = v.Item(1)
    Set v = stack.Item(ss + 3).value
    For ix = 1 To v.Count
        Set kv = v.Item(ix)
        kv.key = key & "." & kv.key
        col.add kv
    Next ix
End Sub

Private Sub afterReduceObject(ByRef newTok As Token, ByRef col As Collection)
    Dim ix As Long
    For ix = 1 To col.Count
        newTok.value = col.Item(ix)
    Next ix
End Sub

Private Sub beforeGetLocale(ByRef stack As Collection, ByRef rule As SyntaxRule, ByRef tr As Translation)
    Dim ss As Long, v As Object, ix As Long
    ss = stack.Count - rule.items.Count
    Set v = stack.Item(ss + 3).value
    tr.locale = v.Item(1)
End Sub

Private Sub afterGetLocale(ByRef newTok As Token, tr As Translation)
    newTok.value = tr
End Sub

Private Sub beforeFinish(ByRef stack As Collection, ByRef rule As SyntaxRule, ByRef tr As Translation)
    Dim ss As Long, iy As Long, line As Long, kv As KeyValPair
    ss = stack.Count - rule.items.Count
    If stack.Item(ss + 3).value.Count > 1 Then
        tr.locale = stack.Item(ss + 1).value.Item(1).locale
        Set tr.items = stack.Item(ss + 3).value
        stack.Item(ss + 3).value = Nothing
    End If
    With ActiveWorkbook.ActiveSheet
        .Cells(1, currentCol + 1) = tr.locale
        .Cells(1, currentCol + 1).ColumnWidth = 40
        .Cells(1, currentCol + 1).Font.Bold = True
        .Cells(1, currentCol) = "KEY"
        .Cells(1, currentCol).ColumnWidth = 40
        .Cells(1, currentCol).Font.Bold = True
        For iy = 1 To tr.items.Count
            Set kv = tr.items.Item(iy)
            If kv Is Nothing Then
                Debug.Print "no value for index: " & iy & ", locale=" & tr.locale
            Else
                .Cells(iy + 1, currentCol + 1) = kv.val
                .Cells(iy + 1, currentCol) = kv.key
            End If
        Next iy
    End With
    currentCol = currentCol + 2
    If maxKeys < tr.items.Count Then
        maxKeys = tr.items.Count
    End If
    For iy = tr.items.Count To 1 Step -1
        tr.items.Remove 1
    Next iy
End Sub

Private Sub afterFinish(ByRef newTok As Token, ByRef tr As Translation)
    If tr.locale <> "" Then
        newTok.value = tr
    End If
End Sub


Private Sub beforeFinishAll(ByRef stack As Collection, ByRef rule As SyntaxRule, ByRef col As Collection)
    Dim ss As Long, v As Object, ix As Long, key As String, kv As KeyValPair
    ss = stack.Count - rule.items.Count
    Set v = stack.Item(ss + 1).value
    For ix = 1 To v.Count
        col.add v.Item(ix)
    Next ix
End Sub

Private Sub afterFinishAll(ByRef newTok As Token, ByRef col As Collection)
    Dim ix As Long, iy As Long, keys As New Dictionary, list As Object, tr As Translation, line As Long, kv As KeyValPair, v As Variant
    newTok.value = col
    
    Set list = CreateObject("System.Collections.ArrayList")
    For ix = 1 To col.Count
        Set tr = col.Item(ix)
        Debug.Print "locale: " & tr.locale & ", items=" & tr.items.Count
        For iy = 1 To tr.items.Count
            Set kv = tr.items.Item(iy)
            If Not keys.Exists(kv.key) Then
                keys.add kv.key, 0
            End If
        Next iy
    Next ix
    
    For Each v In keys.keys
        list.add v
    Next v
    
    keys.RemoveAll
    
    list.sort
    
    line = 2
    For Each v In list
        keys.add v, line
        line = line + 1
    Next v
    
    ActiveWorkbook.ActiveSheet.Cells.Clear
    
    For ix = 1 To col.Count
        Set tr = col.Item(ix)
        With ActiveWorkbook.ActiveSheet
            .Cells(1, ix + 1) = tr.locale
            .Cells(1, ix + 1).ColumnWidth = 40
            .Cells(1, ix + 1).Font.Bold = True
            .Cells(1, 1) = "KEY"
            .Cells(1, 1).ColumnWidth = 40
            .Cells(1, 1).Font.Bold = True
        End With
        For iy = 1 To tr.items.Count
            Set kv = tr.items.Item(iy)
            line = keys.Item(kv.key)
            With ActiveWorkbook.ActiveSheet
                .Cells(line, ix + 1) = kv.val
                .Cells(line, 1) = kv.key
            End With
        Next iy
    Next ix
End Sub

