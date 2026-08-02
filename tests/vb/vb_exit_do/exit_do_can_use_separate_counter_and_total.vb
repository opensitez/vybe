' vybe-test: vb/vb_exit_do/exit_do_can_use_separate_counter_and_total
' origin: languages/vb/tests/vb/test_vb_exit_do.rs

Module M
    Sub Main()
        Dim count As Integer = 0
        Dim total As Integer = 0
        Do
            count = count + 1
            total = total + count
            If count = 4 Then Exit Do
        Loop
        Console.WriteLine(total)
    End Sub
End Module
