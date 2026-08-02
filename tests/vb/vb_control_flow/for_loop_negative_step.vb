' vybe-test: vb/vb_control_flow/for_loop_negative_step
' origin: languages/vb/tests/vb/test_vb_control_flow.rs

Module M
    Sub Main()
        For i As Integer = 5 To 1 Step -1
            Console.WriteLine(i)
        Next
    End Sub
End Module
