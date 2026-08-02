' vybe-test: vb/vb_control_flow_edges/select_case_range
' origin: languages/vb/tests/vb/test_vb_control_flow_edges.rs

Module M
    Sub Main()
        Dim val = 5
        Select Case val
            Case 1 To 10
                Console.WriteLine("1-10")
            Case Else
                Console.WriteLine("Other")
        End Select
    End Sub
End Module
