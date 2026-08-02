' vybe-test: vb/vb_control_flow_edges/select_case_comma
' origin: languages/vb/tests/vb/test_vb_control_flow_edges.rs

Module M
    Sub Main()
        Dim val = 2
        Select Case val
            Case 1, 2, 3
                Console.WriteLine("Matched")
            Case Else
                Console.WriteLine("Other")
        End Select
    End Sub
End Module
