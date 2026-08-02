' vybe-test: vb/vb_do_loop_while_until_variants/test_vb_do_loop_with_return_statement_inside
' origin: languages/vb/tests/vb/test_vb_do_loop_while_until_variants.rs

Module Program
    Private Function FindTarget() As Integer
        Dim i = 0
        Do
            i += 1
            If i = 5 Then Return i * 10
        Loop While True
        Return -1
    End Function

    Sub Main()
        Console.WriteLine(FindTarget())
    End Sub
End Module
