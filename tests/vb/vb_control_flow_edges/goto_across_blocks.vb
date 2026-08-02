' vybe-test: vb/vb_control_flow_edges/goto_across_blocks
' origin: languages/vb/tests/vb/test_vb_control_flow_edges.rs

Module M
    Sub Main()
        If True Then
            GoTo Target
        End If
        
        Console.WriteLine("Skipped")
        
    Target:
        Console.WriteLine("Target")
    End Sub
End Module
