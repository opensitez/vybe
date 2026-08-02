' vybe-test: vb/vb_exit_do/exit_do_skips_tail_work_after_break
' origin: languages/vb/tests/vb/test_vb_exit_do.rs

Module M
    Sub Main()
        Dim total As Integer = 0
        Do
            total = total + 1
            If total >= 2 Then Exit Do
            total = total + 10
        Loop
        Console.WriteLine(total)
    End Sub
End Module
