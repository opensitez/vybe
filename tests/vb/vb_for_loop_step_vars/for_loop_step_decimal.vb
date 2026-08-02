' vybe-test: vb/vb_for_loop_step_vars/for_loop_step_decimal
' origin: languages/vb/tests/vb/test_vb_for_loop_step_vars.rs

Module M
    Sub Main()
        ' Using Decimal for exact step
        For i As Decimal = 0D To 1D Step 0.5D
            Console.WriteLine(i)
        Next
    End Sub
End Module
