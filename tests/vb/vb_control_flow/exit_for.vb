' vybe-test: vb/vb_control_flow/exit_for
' origin: languages/vb/tests/vb/test_vb_control_flow.rs

Module M
    Sub Main()
        For i As Integer = 1 To 100
            If i > 3 Then Exit For
            Console.WriteLine(i)
        Next
    End Sub
End Module
