' vybe-test: vb/vb_do_loop_while_until_variants/test_vb_do_loop_while_bottom_condition
' origin: languages/vb/tests/vb/test_vb_do_loop_while_until_variants.rs

Module Program
    Sub Main()
        Dim count = 0
        Do
            count += 1
        Loop While count < 3
        Console.WriteLine(count)
    End Sub
End Module
