' vybe-test: vb/vb_exit_do/exit_do_leaves_accumulator_available_after_loop
' origin: languages/vb/tests/vb/test_vb_exit_do.rs

Module M
    Sub Main()
        Dim total As Integer = 1
        Do
            total = total * 2
            If total >= 8 Then Exit Do
        Loop
        Console.WriteLine(total + 1)
    End Sub
End Module
