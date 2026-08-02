' vybe-test: vb/vb_exit_do/exit_do_can_overshoot_threshold_before_stopping
' origin: languages/vb/tests/vb/test_vb_exit_do.rs

Module M
    Sub Main()
        Dim total As Integer = 0
        Do
            total = total + 4
            If total >= 7 Then Exit Do
        Loop
        Console.WriteLine(total)
    End Sub
End Module
