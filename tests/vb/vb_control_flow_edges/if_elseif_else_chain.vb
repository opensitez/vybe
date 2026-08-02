' vybe-test: vb/vb_control_flow_edges/if_elseif_else_chain
' origin: languages/vb/tests/vb/test_vb_control_flow_edges.rs

Module M
    Sub Main()
        Dim val = 3
        If val = 1 Then
            Console.WriteLine("1")
        ElseIf val = 2 Then
            Console.WriteLine("2")
        ElseIf val = 3 Then
            Console.WriteLine("3")
        Else
            Console.WriteLine("Other")
        End If
    End Sub
End Module
