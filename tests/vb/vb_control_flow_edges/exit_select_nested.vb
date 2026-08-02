' vybe-test: vb/vb_control_flow_edges/exit_select_nested
' origin: languages/vb/tests/vb/test_vb_control_flow_edges.rs

Module M
    Sub Main()
        Dim val = 1
        Select Case val
            Case 1
                For i = 1 To 5
                    If i = 3 Then Exit Select
                    Console.WriteLine(i)
                Next
            Case 2
                Console.WriteLine("Two")
        End Select
        Console.WriteLine("Done")
    End Sub
End Module
