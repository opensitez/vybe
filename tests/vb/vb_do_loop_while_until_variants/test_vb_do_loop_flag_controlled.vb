' vybe-test: vb/vb_do_loop_while_until_variants/test_vb_do_loop_flag_controlled
' origin: languages/vb/tests/vb/test_vb_do_loop_while_until_variants.rs

Module Program
    Sub Main()
        Dim running = True
        Dim stepCount = 0
        Do While running
            stepCount += 1
            If stepCount = 4 Then running = False
        Loop
        Console.WriteLine(stepCount)
    End Sub
End Module
