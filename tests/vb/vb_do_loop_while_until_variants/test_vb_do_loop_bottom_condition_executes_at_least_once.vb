' vybe-test: vb/vb_do_loop_while_until_variants/test_vb_do_loop_bottom_condition_executes_at_least_once
' origin: languages/vb/tests/vb/test_vb_do_loop_while_until_variants.rs

Module Program
    Sub Main()
        Dim count = 10
        Do
            count += 1
        Loop While count < 5 ' False on first check, but loop body ran once!
        Console.WriteLine(count)
    End Sub
End Module
