' vybe-test: vb/vb_control_flow/nested_loops
' origin: languages/vb/tests/vb/test_vb_control_flow.rs

Module M
    Sub Main()
        Dim count As Integer = 0
        For i As Integer = 1 To 3
            For j As Integer = 1 To 3
                count = count + 1
            Next
        Next
        Console.WriteLine(count)
    End Sub
End Module
