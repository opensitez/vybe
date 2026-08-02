' vybe-test: vb/vb_control_flow/for_loop_step
' origin: languages/vb/tests/vb/test_vb_control_flow.rs

Module M
    Sub Main()
        Dim sum As Integer = 0
        For i As Integer = 0 To 10 Step 2
            sum = sum + i
        Next
        Console.WriteLine(sum)
    End Sub
End Module
