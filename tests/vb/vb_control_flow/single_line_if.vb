' vybe-test: vb/vb_control_flow/single_line_if
' origin: languages/vb/tests/vb/test_vb_control_flow.rs

Module M
    Sub Main()
        Dim x As Integer = 5
        If x > 3 Then Console.WriteLine("big") Else Console.WriteLine("small")
    End Sub
End Module
