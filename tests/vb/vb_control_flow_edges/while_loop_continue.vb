' vybe-test: vb/vb_control_flow_edges/while_loop_continue
' origin: languages/vb/tests/vb/test_vb_control_flow_edges.rs

Module M
    Sub Main()
        Dim i = 0
        Dim sum = 0
        While i < 3
            i += 1
            If i = 2 Then Continue While
            sum += i
        End While
        Console.WriteLine(sum) ' 1 + 3 = 4
    End Sub
End Module
