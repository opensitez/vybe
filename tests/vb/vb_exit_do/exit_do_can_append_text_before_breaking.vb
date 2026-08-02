' vybe-test: vb/vb_exit_do/exit_do_can_append_text_before_breaking
' origin: languages/vb/tests/vb/test_vb_exit_do.rs

Module M
    Sub Main()
        Dim text As String = ""
        Dim count As Integer = 0
        Do
            count = count + 1
            text = text & count
            If count = 3 Then Exit Do
        Loop
        Console.WriteLine(text)
    End Sub
End Module
