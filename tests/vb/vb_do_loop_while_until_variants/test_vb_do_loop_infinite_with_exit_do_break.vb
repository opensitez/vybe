' vybe-test: vb/vb_do_loop_while_until_variants/test_vb_do_loop_infinite_with_exit_do_break
' origin: languages/vb/tests/vb/test_vb_do_loop_while_until_variants.rs

Module Program
    Sub Main()
        Dim i = 0
        Do
            i += 5
            If i > 20 Then Exit Do
        Loop
        Console.WriteLine(i)
    End Sub
End Module
