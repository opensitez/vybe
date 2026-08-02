' vybe-test: vb/vb_do_loop_while_until_variants/test_vb_do_until_complex_boolean_expression
' origin: languages/vb/tests/vb/test_vb_do_loop_while_until_variants.rs

Module Program
    Sub Main()
        Dim x = 0
        Do Until x >= 10 OrElse x = 5
            x += 1
        Loop
        Console.WriteLine(x)
    End Sub
End Module
