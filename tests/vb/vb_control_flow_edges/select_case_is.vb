' vybe-test: vb/vb_control_flow_edges/select_case_is
' origin: languages/vb/tests/vb/test_vb_control_flow_edges.rs

Module M
    Sub Main()
        Dim val = 15
        Select Case val
            Case Is > 20
                Console.WriteLine(">20")
            Case Is > 10
                Console.WriteLine(">10")
        End Select
    End Sub
End Module
