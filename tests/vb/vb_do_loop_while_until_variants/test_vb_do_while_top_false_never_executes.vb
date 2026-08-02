' vybe-test: vb/vb_do_loop_while_until_variants/test_vb_do_while_top_false_never_executes
' origin: languages/vb/tests/vb/test_vb_do_loop_while_until_variants.rs

Module Program
    Sub Main()
        Dim count = 10
        Do While count < 5 ' False initially
            count += 1
        Loop
        Console.WriteLine(count)
    End Sub
End Module
