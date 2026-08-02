' vybe-test: vb/vb_control_flow/while_loop
' origin: languages/vb/tests/vb/test_vb_control_flow.rs

Module M
    Sub Main()
        Dim i As Integer = 0
        While i < 3
            Console.WriteLine(i)
            i = i + 1
        End While
    End Sub
End Module
