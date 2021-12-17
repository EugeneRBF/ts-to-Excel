Attribute VB_Name = "tsSave"
Option Explicit

Dim localeChars As String, latinChars As String

Public Sub Menu_SaveAsTS_OnAction()
    Dim row As Long, col As Long, keyCol As Long, keyRow As Long, level As Long, locale As String, key As String, val As String
    Dim kv As KeyValPair, dict As Dictionary, tok As Dictionary, list As Collection, tabstop As Long, k As Variant, i As Variant
    
    Dim myDir As String, stream As New ADODB.stream
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        .Title = "”кажите директорию в которой сохран€ть файлы"
        
 
        If .Show <> 0 Then
            myDir = .SelectedItems(1)
        Else
            Exit Sub
        End If
     End With
    
    localeChars = "abcdefghijklmnopqrstuvwxyz"
    latinChars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ_0123456789"
    keyRow = 1
    keyCol = 1
    col = 2
    tabstop = 3
    level = 1
    Set dict = New Dictionary
    
    With ActiveSheet
        Do
            locale = .Cells(keyRow, col)
            If locale <> "" And isValidLocale(locale) Then
                dict.RemoveAll
                
                ' Load each selected file
                stream.Charset = "utf-8"
                stream.Open
                
                row = keyRow + 1
                While .Cells(row, keyCol) <> ""
                    key = .Cells(row, keyCol)
                    val = .Cells(row, col)
                    addRow key, val, locale, dict
                    row = row + 1
                Wend
                
                checkAndMergeKeys dict
        
                Debug.Print beforeEach(locale) & "{"
                stream.WriteText beforeEach(locale) & "{" & vbNewLine
                
                For Each k In dict.keys
                    If k = "." Then
                        Set list = dict.Item(k)
                        For Each i In list
                            Debug.Print offset(level, tabstop) & quoteStr(i.key) & ": " & quoteStr(i.val, True) & ","
                            stream.WriteText offset(level, tabstop) & quoteStr(i.key) & ": " & quoteStr(i.val, True) & "," & vbNewLine
                        Next i
                    Else
                        Set tok = dict.Item(k)
                        printDict stream, tok, level, tabstop, k
                    End If
                Next k
                
                Debug.Print "}"
                stream.WriteText "}" & vbNewLine
                stream.SaveToFile myDir & "\" & "translations-new." & locale & ".ts", adSaveCreateOverWrite
                stream.Close
            End If
            col = col + 1
        Loop While locale <> ""
    End With
End Sub

Private Sub checkAndMergeKeys(ByRef dict As Dictionary)
    Dim list As Collection, added As Collection, k As Variant, nested As Collection, kv As KeyValPair, ix, iy As Long, tok As Dictionary
    Set list = dict.Item(".")
    For Each k In dict.keys
        If k = "" Then
            Set tok = dict.Item(k)
            checkAndMergeKeys tok
            Set added = New Collection
            If tok.Exists(".") Then
                Set nested = tok.Item(".")
                For iy = 1 To nested.Count
                    Set kv = New KeyValPair
                    kv.key = k & "." & nested.Item(iy).key
                    kv.val = nested.Item(iy).val
                    added.add kv
                Next iy
                If dict.Exists(k) Then dict.Remove k
            End If
            If added.Count > 0 Then sortAndMerge list, added
        ElseIf k <> "." Then
            Set tok = dict.Item(k)
            checkAndMergeKeys tok
            Set added = New Collection
            For ix = 1 To list.Count
                If list.Item(ix).key = k Then
                    If tok.Exists(".") Then
                        Set nested = tok.Item(".")
                        For iy = 1 To nested.Count
                            Set kv = New KeyValPair
                            kv.key = k & "." & nested.Item(iy).key
                            kv.val = nested.Item(iy).val
                            added.add kv
                        Next iy
                        
                        If dict.Exists(k) Then dict.Remove k
                    End If
                End If
            Next ix
            If added.Count > 0 Then sortAndMerge list, added
        End If
    Next k
End Sub

Private Sub sortAndMerge(ByRef col As Collection, ByRef added As Collection)
    Dim arr As Object, ix As Long, kv As KeyValPair, dict As New Dictionary, k As Variant
    Set arr = CreateObject("System.Collections.ArrayList")
    For ix = 1 To added.Count
        Set kv = added.Item(ix)
        col.add kv
    Next ix
    For ix = 1 To col.Count
        Set kv = col.Item(ix)
        dict.add kv.key, kv
        arr.add kv.key
    Next ix
    arr.sort
    For ix = col.Count To 1 Step -1
        col.Remove ix
    Next ix
    For Each k In arr
        col.add dict.Item(k)
    Next k
End Sub

Private Sub printDict(stream As ADODB.stream, dict As Dictionary, level As Long, tabstop As Long, k As Variant)
    Dim tok As Dictionary, list As Collection, i As Variant
    Debug.Print offset(level, tabstop) & quoteStr(k) & ": {"
    stream.WriteText offset(level, tabstop) & quoteStr(k) & ": {" & vbNewLine
    For Each k In dict.keys
        If k = "." Then
            Set list = dict.Item(k)
            For Each i In list
                Debug.Print offset(level + 1, tabstop) & quoteStr(i.key) & ": " & quoteStr(i.val, True) & ","
                stream.WriteText offset(level + 1, tabstop) & quoteStr(i.key) & ": " & quoteStr(i.val, True) & "," & vbNewLine
            Next i
        Else
            Set tok = dict.Item(k)
            printDict stream, tok, level + 1, tabstop, k
        End If
    Next k
    Debug.Print offset(level, tabstop) & "},"
    stream.WriteText offset(level, tabstop) & "}," & vbNewLine
End Sub

Private Function quoteStr(val As Variant, Optional asVal As Boolean = False) As Variant
    Dim escaped As Variant, ix As Long
    If VBA.InStr(1, val, "'") > 0 Then
        If VBA.InStr(1, val, """") > 0 Then
            escaped = ""
            For ix = 1 To VBA.Len(val)
                If VBA.Mid$(val, ix, 1) = "'" Then
                    escaped = escaped & "\"
                End If
                escaped = escaped & VBA.Mid$(val, ix, 1)
            Next ix
            quoteStr = "'" & escaped & "'"
        Else
            quoteStr = """" & val & """"
        End If
    ElseIf VBA.InStr(1, val, "`") > 0 Then
        If VBA.InStr(1, val, "'") > 0 Then
            escaped = ""
            For ix = 1 To VBA.Len(val)
                If VBA.Mid$(val, ix, 1) = "'" Then
                    escaped = escaped & "\"
                End If
                escaped = escaped & VBA.Mid$(val, ix, 1)
            Next ix
            quoteStr = "'" & escaped & "'"
        Else
            quoteStr = "'" & val & "'"
        End If
    ElseIf VBA.InStr(1, val, """") > 0 Then
        If VBA.InStr(1, val, "'") > 0 Then
            escaped = ""
            For ix = 1 To VBA.Len(val)
                If VBA.Mid$(val, ix, 1) = "'" Then
                    escaped = escaped & "\"
                End If
                escaped = escaped & VBA.Mid$(val, ix, 1)
            Next ix
            quoteStr = "'" & escaped & "'"
        Else
            quoteStr = "'" & val & "'"
        End If
    ElseIf asVal Or VBA.InStr(1, "0123456789_", VBA.Mid$(val, 1, 1)) > 0 Or Not isLatinOnly(val) Then
        quoteStr = "'" & val & "'"
    Else
        quoteStr = val
    End If
End Function

Private Function isLatinOnly(val As Variant) As Boolean
    Dim ix As Long
    isLatinOnly = True
    For ix = 1 To VBA.Len(val)
        If VBA.InStr(1, latinChars, VBA.Mid$(val, ix, 1), vbBinaryCompare) < 1 Then
            isLatinOnly = False
            Exit For
        End If
    Next ix
End Function

Private Function beforeAll() As String
    beforeAll = ""
End Function

Private Function beforeEach(lo As String) As String
    beforeEach = "import {translations} from './translations';" & vbNewLine _
        & "translations['" & lo & "'] = "
End Function

Private Function offset(ByRef level As Long, ByRef tabstop As Long) As String
    offset = VBA.Space$(tabstop * level)
End Function

Private Function isValidLocale(lo As String) As Boolean
    Dim ch As Integer
    isValidLocale = False
    If VBA.Len(lo) <> 2 Then
        Exit Function
    End If
    isValidLocale = True
    For ch = 1 To VBA.Len(lo)
        If VBA.InStr(1, localeChars, VBA.Mid$(lo, ch, 1), vbBinaryCompare) < 1 Then
            isValidLocale = False
            Exit For
        End If
    Next ch
End Function

Private Sub addRow(ByRef key As String, ByRef val As String, ByRef locale As String, dict As Dictionary)
    Dim v() As String, ix As Long, k As String, tok As Dictionary, parentTok As Dictionary, lastKey As String, col As Collection
    Dim kv As KeyValPair, obj As Variant
    v = VBA.Split(key, ".")
    Set parentTok = dict
    For ix = LBound(v) To UBound(v) - 1
        If parentTok.Exists(v(ix)) Then
            Set tok = parentTok.Item(v(ix))
        Else
            Set tok = New Dictionary
            parentTok.add v(ix), tok
        End If
        Set parentTok = tok
        lastKey = v(ix)
    Next ix
    If Not parentTok.Exists(".") Then
        Set col = New Collection
        parentTok.add ".", col
    Else
        Set col = parentTok.Item(".")
    End If
    Set kv = New KeyValPair
    kv.key = v(UBound(v))
    kv.val = val
    col.add kv ', v(UBound(v))
End Sub
