' vybe-test: vb/vb_control_flow/exit_sub
' origin: languages/vb/tests/vb/test_vb_control_flow.rs

Module M
    Sub Process(x As Integer)
        If x < 0 Then Exit Sub
        Console.WriteLine(x)
    End Sub
    Sub Main()
        Process(5)
        Process(-1)
        Process(10)
    End Sub
End Module
