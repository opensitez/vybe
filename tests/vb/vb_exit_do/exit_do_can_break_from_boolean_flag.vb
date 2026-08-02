' vybe-test: vb/vb_exit_do/exit_do_can_break_from_boolean_flag
' origin: languages/vb/tests/vb/test_vb_exit_do.rs

Module M
    Sub Main()
        Dim count As Integer = 0
        Dim shouldStop As Boolean = False
        Do
            count = count + 1
            shouldStop = count >= 3
            If shouldStop Then Exit Do
        Loop
        Console.WriteLine(count)
    End Sub
End Module
