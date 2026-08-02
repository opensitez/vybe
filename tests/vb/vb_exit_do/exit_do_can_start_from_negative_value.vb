' vybe-test: vb/vb_exit_do/exit_do_can_start_from_negative_value
' origin: languages/vb/tests/vb/test_vb_exit_do.rs

Module M
    Sub Main()
        Dim total As Integer = -2
        Do
            total = total + 3
            If total >= 4 Then Exit Do
        Loop
        Console.WriteLine(total)
    End Sub
End Module
