' vybe-test: vb/vb_do_loop_while_until_variants/test_vb_do_until_bottom_condition_executes_body_once
' origin: languages/vb/tests/vb/test_vb_do_loop_while_until_variants.rs

Module Program
    Sub Main()
        Dim count = 100
        Do
            count += 5
        Loop Until count > 50 ' True after first check, loop ends!
        Console.WriteLine(count)
    End Sub
End Module
