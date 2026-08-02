' vybe-test: vb/vb_exit_do/exit_do_can_break_after_even_iteration
' origin: languages/vb/tests/vb/test_vb_exit_do.rs

Module M
    Sub Main()
        Dim count As Integer = 0
        Do
            count = count + 1
            If count Mod 2 = 0 Then
                Exit Do
            End If
        Loop
        Console.WriteLine(count)
    End Sub
End Module
