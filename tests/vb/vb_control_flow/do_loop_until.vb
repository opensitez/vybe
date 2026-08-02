' vybe-test: vb/vb_control_flow/do_loop_until
' origin: languages/vb/tests/vb/test_vb_control_flow.rs

Module M
    Sub Main()
        Dim i As Integer = 0
        Do
            i = i + 1
        Loop Until i >= 3
        Console.WriteLine(i)
    End Sub
End Module
