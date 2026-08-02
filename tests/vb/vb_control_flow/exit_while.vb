' vybe-test: vb/vb_control_flow/exit_while
' origin: languages/vb/tests/vb/test_vb_control_flow.rs

Module M
    Sub Main()
        Dim i As Integer = 0
        While True
            i = i + 1
            If i = 3 Then Exit While
        End While
        Console.WriteLine(i)
    End Sub
End Module
