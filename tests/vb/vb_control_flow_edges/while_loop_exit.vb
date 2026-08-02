' vybe-test: vb/vb_control_flow_edges/while_loop_exit
' origin: languages/vb/tests/vb/test_vb_control_flow_edges.rs

Module M
    Sub Main()
        Dim i = 0
        While i < 10
            If i = 2 Then Exit While
            i += 1
        End While
        Console.WriteLine(i)
    End Sub
End Module
