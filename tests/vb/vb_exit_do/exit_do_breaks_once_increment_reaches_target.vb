' vybe-test: vb/vb_exit_do/exit_do_breaks_once_increment_reaches_target
' origin: languages/vb/tests/vb/test_vb_exit_do.rs

Module M
    Sub Main()
        Dim total As Integer = 0
        Do
            total = total + 1
            If total >= 3 Then Exit Do
        Loop
        Console.WriteLine(total)
    End Sub
End Module
