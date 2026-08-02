' vybe-test: vb/vb_for_loop_step_vars/for_loop_modification_during_loop
' origin: languages/vb/tests/vb/test_vb_for_loop_step_vars.rs

Module M
    Sub Main()
        Dim endVal = 3
        For i As Integer = 1 To endVal
            Console.WriteLine(i)
            endVal = 10 ' modifying endVal has no effect on the loop condition
        Next
    End Sub
End Module
