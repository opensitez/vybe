' vybe-test: vb/vb_control_flow_edges/for_loop_step_negative
' origin: languages/vb/tests/vb/test_vb_control_flow_edges.rs

Module M
    Sub Main()
        For i = 3 To 1 Step -1
            Console.WriteLine(i)
        Next
    End Sub
End Module
