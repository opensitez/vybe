' vybe-test: vb/vb_control_flow_edges/for_loop_step_decimal
' origin: languages/vb/tests/vb/test_vb_control_flow_edges.rs

Module M
    Sub Main()
        For i As Decimal = 1.5D To 2.5D Step 0.5D
            Console.WriteLine(i)
        Next
    End Sub
End Module
