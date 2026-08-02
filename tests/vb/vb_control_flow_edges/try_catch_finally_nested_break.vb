' vybe-test: vb/vb_control_flow_edges/try_catch_finally_nested_break
' origin: languages/vb/tests/vb/test_vb_control_flow_edges.rs

Module M
    Sub Main()
        For i = 1 To 3
            Try
                If i = 2 Then Exit For
            Finally
                Console.WriteLine("Finally" & i)
            End Try
        Next
        Console.WriteLine("Done")
    End Sub
End Module
