' vybe-test: vb/vb_do_loop_while_until_variants/test_vb_do_loop_modifying_step_variable
' origin: languages/vb/tests/vb/test_vb_do_loop_while_until_variants.rs

Module Program
    Sub Main()
        Dim val = 1
        Do While val < 100
            val *= 2
        Loop
        Console.WriteLine(val)
    End Sub
End Module
