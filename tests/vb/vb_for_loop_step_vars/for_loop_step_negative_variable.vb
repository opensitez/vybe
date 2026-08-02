' vybe-test: vb/vb_for_loop_step_vars/for_loop_step_negative_variable
' origin: languages/vb/tests/vb/test_vb_for_loop_step_vars.rs

Module M
    Sub Main()
        Dim startVal = 5
        Dim endVal = 1
        Dim stepVal = -2
        
        For i As Integer = startVal To endVal Step stepVal
            Console.WriteLine(i)
        Next
    End Sub
End Module
