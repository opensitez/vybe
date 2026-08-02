' vybe-test: vb/vb_control_flow/do_while_loop
' origin: languages/vb/tests/vb/test_vb_control_flow.rs

Module M
    Sub Main()
        Dim i As Integer = 0
        Do While i < 3
            Console.WriteLine(i)
            i = i + 1
        Loop
    End Sub
End Module
