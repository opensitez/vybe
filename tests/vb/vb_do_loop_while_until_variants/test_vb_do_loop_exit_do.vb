' vybe-test: vb/vb_do_loop_while_until_variants/test_vb_do_loop_exit_do
' origin: languages/vb/tests/vb/test_vb_do_loop_while_until_variants.rs

Module Program
    Sub Main()
        Dim count = 0
        Do
            count += 1
            If count = 3 Then Exit Do
        Loop While count < 10
        Console.WriteLine(count)
    End Sub
End Module
