Public Class Converter
    Private WithEvents parser As New TsParse
    Private Sub ParserBtn_Click(sender As Object, e As EventArgs) Handles parserBtn.Click
        parser.doParse()
    End Sub

    Private Sub parser_NewLang(lang As String) Handles parser.NewLang
        Console.WriteLine("newLang {0}", lang)
        LangsList.Items.Add(lang)
    End Sub

    Private Sub parser_Finish(ByRef content As Collection) Handles parser.Finish
        Console.WriteLine("Finish {0}", content)
        For Each l In LangsList.Items
            For Each tr In content
                If l.Equals(tr.locale) Then
                    l = l & ", items:" & tr.items.count
                End If
            Next
        Next
    End Sub
End Class
