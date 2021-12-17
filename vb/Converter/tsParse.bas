'
' (C) 2021 Eugene Torkhov
'
Option Explicit On

Imports Microsoft.Office.Interop

Public Class TsParse
    Dim rules As Collection
    Dim objApp As Excel.Application
    Dim objBook As Excel._Workbook

    Public Event NewLang(ByVal lang As String)
    Public Event Finish(ByRef content As Collection)

    Public Sub DoParse()
        Dim stream As New ADODB.Stream, text As String
        Dim parser As New Lexem, tkn As New Token, stack As New Collection, dialog As OpenFileDialog

        Dim my As New TsParse

        my.prepareRules()

        dialog = New OpenFileDialog
        text = ""
        With dialog
            .Multiselect = True

            If .ShowDialog() = DialogResult.OK Then
                For Each file In .FileNames
                    ' Load each selected file
                    stream.Charset = "utf-8"
                    stream.Open()
                    stream.LoadFromFile(file)
                    text = text & stream.ReadText() & vbNewLine
                    stream.Close()
                Next
            End If
        End With

        If Len(text) > 0 Then

            parser.init(text)

            Do
                parser.parse()
                'Console.WriteLine(parser.index, parser.toStr())
                tkn = tkn.toToken(parser)
                stack.Add(Item:=tkn)
                my.tryReduce(stack)
            Loop While parser.lexemType <> LexemType.EOF 'And jsonIndex < VBA.Len(text)

            stack.Clear()
            'MsgBox("Complete")
        End If

    End Sub

    Private Sub PrepareRules()
        Dim rule As New SyntaxRule

        rules = New Collection

        rules.Add(Item:=rule.create(Action.FinishAll, TokenType.Programm).add(TokenType.StatementList).add(TokenType.EOP))
        rules.Add(Item:=rule.create(Action.Copy, TokenType.Statement).add(TokenType.Statement).add(TokenType.tSemiColon))             ' stmt ;
        rules.Add(Item:=rule.create(Action.None, TokenType.Statement).add(TokenType.tIdentifier).add(TokenType.tAssign).add(TokenType.vString)) ' ident = 'string'
        rules.Add(Item:=rule.create(Action.None, TokenType.Statement).add(TokenType.tImport).add(TokenType.vString))                  ' import 'string'
        rules.Add(Item:=rule.create(Action.None, TokenType.Statement).add(TokenType.tImport).add(TokenType.tCurlyStart).add(TokenType.tIdentifier).add(TokenType.tCurlyEnd).add(TokenType.tIdentifier).add(TokenType.vString)) ' import { class } from 'string'
        rules.Add(Item:=rule.create(Action.None, TokenType.Statement).add(TokenType.tImport).add(TokenType.tCurlyStart).add(TokenType.tIdentList).add(TokenType.tCurlyEnd).add(TokenType.tIdentifier).add(TokenType.vString)) ' import { class } from 'string'
        rules.Add(Item:=rule.create(Action.None, TokenType.Statement).add(TokenType.tImport).add(TokenType.tIdentifier).add(TokenType.vString)) ' import * from 'string'
        rules.Add(Item:=rule.create(Action.None, TokenType.tIdentList).add(TokenType.tIdentifier).add(TokenType.tComma).add(TokenType.tIdentifier))     ' identList := <ident>, <ident> ;
        rules.Add(Item:=rule.create(Action.None, TokenType.tIdentList).add(TokenType.tIdentList).add(TokenType.tComma).add(TokenType.tIdentifier))     ' identList := <ident>, <ident> ;
        rules.Add(Item:=rule.create(Action.Merge, TokenType.StatementList).add(TokenType.StatementList).add(TokenType.Statement))     ' stmtList := <stmtList> <stmt>
        rules.Add(Item:=rule.create(Action.Copy, TokenType.StatementList).add(TokenType.StatementList).add(TokenType.tSemiColon))     ' stmtList := <stmtList> ;
        rules.Add(Item:=rule.create(Action.Copy, TokenType.StatementList).add(TokenType.Statement))                         ' stmtList := <stmt>
        rules.Add(Item:=rule.create(Action.Finish, TokenType.Statement).add(TokenType.tIdentifier).add(TokenType.tAssign).add(TokenType.vObject))   ' stmt := <ident> = <object>
        rules.Add(Item:=rule.create(Action.GetLocale, TokenType.tIdentifier).add(TokenType.tIdentifier).add(TokenType.tBrackedStart).add(TokenType.vString).add(TokenType.tBrackedEnd))   ' ident = <ident> [ 'string' ]
        rules.Add(Item:=rule.create(Action.Copy2, TokenType.tIdentifier).add(TokenType.tIdentifier).add(TokenType.tIdentifier))        ' ident := export const <ident>
        rules.Add(Item:=rule.create(Action.Copy2, TokenType.vObject).add(TokenType.tCurlyStart).add(TokenType.oPropertyList).add(TokenType.tCurlyEnd))
        rules.Add(Item:=rule.create(Action.None, TokenType.vObject).add(TokenType.tCurlyStart).add(TokenType.tCurlyEnd))
        rules.Add(Item:=rule.create(Action.ReduceString, TokenType.oProperty).add(TokenType.tIdentifier).add(TokenType.tColon).add(TokenType.vString))
        rules.Add(Item:=rule.create(Action.ReduceObject, TokenType.oProperty).add(TokenType.tIdentifier).add(TokenType.tColon).add(TokenType.vObject))
        rules.Add(Item:=rule.create(Action.ReduceString, TokenType.oProperty).add(TokenType.vString).add(TokenType.tColon).add(TokenType.vString))
        rules.Add(Item:=rule.create(Action.ReduceObject, TokenType.oProperty).add(TokenType.vString).add(TokenType.tColon).add(TokenType.vObject))
        rules.Add(Item:=rule.create(Action.Copy, TokenType.oPropertyList).add(TokenType.oPropertyList).add(TokenType.tComma))
        rules.Add(Item:=rule.create(Action.Merge, TokenType.oPropertyList).add(TokenType.oPropertyList).add(TokenType.oProperty))
        rules.Add(Item:=rule.create(Action.Copy, TokenType.oPropertyList).add(TokenType.oProperty))
        rules.Add(Item:=rule.create(Action.Copy, TokenType.vString).add(TokenType.vString).add(TokenType.tComment))
        rules.Add(Item:=rule.create(Action.None, TokenType.tComma).add(TokenType.tComma).add(TokenType.tComment))
        rules.Add(Item:=rule.create(Action.Copy, TokenType.vObject).add(TokenType.vObject).add(TokenType.tComment))
        rules.Add(Item:=rule.create(Action.None, TokenType.tAssign).add(TokenType.tAssign).add(TokenType.tComment))
        rules.Add(Item:=rule.create(Action.None, TokenType.tBrackedStart).add(TokenType.tBrackedStart).add(TokenType.tComment))
        rules.Add(Item:=rule.create(Action.None, TokenType.tBrackedEnd).add(TokenType.tBrackedEnd).add(TokenType.tComment))
        rules.Add(Item:=rule.create(Action.None, TokenType.tCurlyStart).add(TokenType.tCurlyStart).add(TokenType.tComment))
        rules.Add(Item:=rule.create(Action.None, TokenType.tCurlyEnd).add(TokenType.tCurlyEnd).add(TokenType.tComment))
        rules.Add(Item:=rule.create(Action.Copy, TokenType.tIdentifier).add(TokenType.tIdentifier).add(TokenType.tComment))
        rules.Add(Item:=rule.create(Action.Copy, TokenType.Statement).add(TokenType.Statement).add(TokenType.tComment))
        rules.Add(Item:=rule.create(Action.Copy, TokenType.StatementList).add(TokenType.StatementList).add(TokenType.tComment))
        'rules.add item:=rule.create(Action.Copy2, TokenType.StatementList).add(TokenType.tComment).add(TokenType.StatementList)
        rules.Add(Item:=rule.create(Action.None, TokenType.Programm).add(TokenType.Programm).add(TokenType.tComment))
        rules.Add(Item:=rule.create(Action.Copy, TokenType.oProperty).add(TokenType.oProperty).add(TokenType.tComment))
        rules.Add(Item:=rule.create(Action.Copy, TokenType.oPropertyList).add(TokenType.oPropertyList).add(TokenType.tComment))
    End Sub

    Private Sub TryReduce(ByRef stack As Collection)
        Dim ix As Long, rule As SyntaxRule, ss As Long, iy As Integer, matched As Boolean, newTok As Token, t As Token
        Dim key As String, val As String, kv As KeyValPair, col As Collection, tr As Translation, wasReduce As Boolean
        'On Error Resume Next
        Do
            wasReduce = False
            For ix = 1 To rules.Count
                rule = rules.Item(ix)
                If stack.Count >= rule.items.Count Then
                    ss = stack.Count - rule.items.Count
                    matched = True
                    For iy = 1 To rule.items.Count
                        t = stack.Item(ss + iy)
                        If t.type <> rule.items.Item(iy) Then
                            matched = False
                            Exit For
                        End If
                    Next iy
                    If matched Then
                        col = New Collection
                        kv = New KeyValPair
                        tr = New Translation
                        Select Case rule.canDo
                            Case Action.ReduceString : beforeReduceString(stack, rule, kv)
                            Case Action.ReduceObject : beforeReduceObject(stack, rule, col)
                            Case Action.Copy : beforeCopy(stack, rule, col)
                            Case Action.Copy2 : beforeCopy2(stack, rule, col)
                            Case Action.Merge : beforeMerge(stack, rule, col)
                            Case Action.GetLocale : beforeGetLocale(stack, rule, tr)
                            Case Action.Finish : beforeFinish(stack, rule, tr)
                            Case Action.FinishAll : beforeFinishAll(stack, rule, col)
                        End Select

                        'Console.WriteLine("Stack reduced to " & rule.toStr)
                        For iy = stack.Count To ss + 1 Step -1
                            stack.Remove(iy)
                        Next iy

                        newTok = New Token With {.type = rule.target}
                        stack.Add(newTok)

                        wasReduce = True

                        Select Case rule.canDo
                            Case Action.ReduceString : afterReduceString(newTok, kv)
                            Case Action.ReduceObject : afterReduceObject(newTok, col)
                            Case Action.Copy : afterCopy(newTok, col)
                            Case Action.Copy2 : afterCopy(newTok, col)
                            Case Action.Merge : afterMerge(newTok, col)
                            Case Action.GetLocale : afterGetLocale(newTok, tr)
                            Case Action.Finish : afterFinish(newTok, tr)
                            Case Action.FinishAll : afterFinishAll(newTok, col)
                        End Select

                        col = Nothing
                        kv = Nothing
                        tr = Nothing
                        'tryReduce stack        'in recursive mode or just leave from the loop
                        Exit For
                    End If
                End If
            Next ix
        Loop While wasReduce
        On Error GoTo 0
    End Sub

    Private Sub BeforeReduceString(ByRef stack As Collection, ByRef rule As SyntaxRule, ByRef kv As KeyValPair)
        Dim ss As Long, tok1 As Token, tok3 As Token
        ss = stack.Count - rule.items.Count
        tok1 = stack.Item(ss + 1)
        tok3 = stack.Item(ss + 3)
        kv.key = tok1.value.Item(1)
        kv.val = tok3.value.Item(1)
    End Sub

    Private Sub AfterReduceString(ByRef newTok As Token, ByRef kv As KeyValPair)
        newTok.value.Add(kv)
    End Sub

    Private Sub BeforeCopy(ByRef stack As Collection, ByRef rule As SyntaxRule, ByRef col As Collection)
        Dim ss As Long, tok As Token, ix As Long
        ss = stack.Count - rule.items.Count
        tok = stack.Item(ss + 1)
        For ix = 1 To tok.value.Count
            col.Add(tok.value.Item(ix))
        Next ix
    End Sub

    Private Sub AfterCopy(ByRef newTok As Token, ByRef col As Collection)
        Dim ix As Long
        For ix = 1 To col.Count
            newTok.value.Add(col.Item(ix))
        Next ix
    End Sub

    Private Sub BeforeCopy2(ByRef stack As Collection, ByRef rule As SyntaxRule, ByRef col As Collection)
        Dim ss As Long, tok As Token, ix As Long
        ss = stack.Count - rule.items.Count
        tok = stack.Item(ss + 2)
        For ix = 1 To tok.value.Count
            col.Add(tok.value.Item(ix))
        Next ix
    End Sub

    Private Sub BeforeMerge(ByRef stack As Collection, ByRef rule As SyntaxRule, ByRef col As Collection)
        Dim ss As Long, tok1 As Token, tok2 As Token, ix As Long
        ss = stack.Count - rule.items.Count
        tok1 = stack.Item(ss + 1)
        For ix = 1 To tok1.value.Count
            col.Add(tok1.value.Item(ix))
        Next ix
        tok2 = stack.Item(ss + 2)
        For ix = 1 To tok2.value.Count
            col.Add(tok2.value.Item(ix))
        Next ix
    End Sub

    Private Sub AfterMerge(ByRef newTok As Token, ByRef col As Collection)
        Dim ix As Long
        For ix = 1 To col.Count
            newTok.value.Add(col.Item(ix))
        Next ix
    End Sub


    Private Sub BeforeReduceObject(ByRef stack As Collection, ByRef rule As SyntaxRule, ByRef col As Collection)
        Dim ss As Long, ix As Long, key As String, kv As KeyValPair, tok As Token
        ss = stack.Count - rule.items.Count
        tok = stack.Item(ss + 1)
        key = tok.value.Item(1)
        tok = stack.Item(ss + 3)
        For ix = 1 To tok.value.Count
            kv = tok.value.Item(ix)
            kv.key = key & "." & kv.key
            col.Add(kv)
        Next ix
    End Sub

    Private Sub AfterReduceObject(ByRef newTok As Token, ByRef col As Collection)
        Dim ix As Long
        For ix = 1 To col.Count
            newTok.value.Add(col.Item(ix))
        Next ix
    End Sub

    Private Sub BeforeGetLocale(ByRef stack As Collection, ByRef rule As SyntaxRule, ByRef tr As Translation)
        Dim ss As Long, v As Token
        ss = stack.Count - rule.items.Count
        v = stack.Item(ss + 3)
        tr.locale = v.value.Item(1)
    End Sub

    Private Sub AfterGetLocale(ByRef newTok As Token, ByVal tr As Translation)
        newTok.value.Add(tr)
        RaiseEvent NewLang(tr.locale)
    End Sub

    Private Sub BeforeFinish(ByRef stack As Collection, ByRef rule As SyntaxRule, ByRef tr As Translation)
        Dim ss As Long, tok As Token, tok1 As Token
        ss = stack.Count - rule.items.Count
        tok = stack.Item(ss + 3)
        If tok.value.Count > 1 Then
            tok1 = stack.Item(ss + 1)
            tr.locale = tok1.value.Item(1).locale
            'tr.locale = tok1.locale
            tr.items = tok.value
        End If
    End Sub

    Private Sub AfterFinish(ByRef newTok As Token, ByRef tr As Translation)
        If tr.locale <> "" Then
            newTok.value.Add(tr)
        End If
    End Sub


    Private Sub BeforeFinishAll(ByRef stack As Collection, ByRef rule As SyntaxRule, ByRef col As Collection)
        Dim ss As Long, v As Token, ix As Long
        ss = stack.Count - rule.items.Count
        v = stack.Item(ss + 1)
        For ix = 1 To v.value.Count
            col.Add(v.value.Item(ix))
        Next ix
    End Sub

    Private Sub AfterFinishAll(ByRef newTok As Token, ByRef col As Collection)
        Dim ix As Long, iy As Long, keys As New Dictionary(Of String, Long), list As New System.Collections.ArrayList, tr As Translation, line As Long, kv As KeyValPair, v As String
        Dim objBooks As Excel.Workbooks
        Dim objSheets As Excel.Sheets
        Dim objSheet As Excel._Worksheet

        RaiseEvent Finish(col)

        objApp = New Excel.Application()
        objBooks = objApp.Workbooks
        objBook = objBooks.Add
        objSheets = objBooks(1).Worksheets
        objSheet = objSheets(1)

        'newTok.value.Add(col)

        For ix = 1 To col.Count
            tr = col.Item(ix)
            Console.WriteLine("locale: " & tr.locale & ", items=" & tr.items.Count)
            For iy = 1 To tr.items.Count
                kv = tr.items.Item(iy)
                If Not keys.TryGetValue(kv.key, line) Then
                    keys.Add(kv.key, 0)
                End If
            Next iy
        Next ix

        For Each v In keys.Keys
            list.Add(v)
        Next v

        keys.Clear()

        list.Sort()

        line = 2
        For Each v In list
            keys.Add(v, line)
            line += 1
        Next v

        objSheet.Cells.Clear()

        For ix = 1 To col.Count
            tr = col.Item(ix)

            With objSheet
                .Cells(1, ix + 1) = tr.locale
                .Cells(1, ix + 1).ColumnWidth = 40
                .Cells(1, ix + 1).Font.Bold = True
                .Cells(1, 1) = "KEY"
                .Cells(1, 1).ColumnWidth = 40
                .Cells(1, 1).Font.Bold = True
            End With
            For iy = 1 To tr.items.Count
                kv = tr.items.Item(iy)
                line = keys.Item(kv.key)
                'Console.WriteLine("key:{0}, val:{1}", kv.key, kv.val)
                With objSheet
                    .Cells(line, ix + 1) = kv.val
                    .Cells(line, 1) = kv.key
                End With
            Next iy
        Next ix
        objApp.Visible = True
        objApp.UserControl = True

        'Clean up a little.
        objSheet = Nothing
        objSheets = Nothing
        objBooks = Nothing
    End Sub

End Class