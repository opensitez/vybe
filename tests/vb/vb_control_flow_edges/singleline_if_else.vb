' vybe-test: vb/vb_control_flow_edges/singleline_if_else
' origin: languages/vb/tests/vb/test_vb_control_flow_edges.rs

Module M
    Sub Main()
        Dim x = 10
        If x > 5 Then Console.WriteLine("Yes") Else Console.WriteLine("No")
    End Sub
End Module
