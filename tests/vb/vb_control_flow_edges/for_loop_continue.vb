' vybe-test: vb/vb_control_flow_edges/for_loop_continue
' origin: languages/vb/tests/vb/test_vb_control_flow_edges.rs

Module M
    Sub Main()
        Dim sum = 0
        For i = 1 To 3
            If i = 2 Then Continue For
            sum += i
        Next
        Console.WriteLine(sum) ' 1 + 3 = 4
    End Sub
End Module
