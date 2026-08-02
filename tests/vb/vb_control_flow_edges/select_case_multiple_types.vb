' vybe-test: vb/vb_control_flow_edges/select_case_multiple_types
' origin: languages/vb/tests/vb/test_vb_control_flow_edges.rs

Option Strict Off
Module M
    Sub Main()
        Dim val As Object = "10"
        
        Select Case val
            Case 10
                Console.WriteLine("Num10")
            Case "10"
                Console.WriteLine("Str10")
        End Select
    End Sub
End Module
