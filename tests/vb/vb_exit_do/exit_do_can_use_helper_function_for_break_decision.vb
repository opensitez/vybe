' vybe-test: vb/vb_exit_do/exit_do_can_use_helper_function_for_break_decision
' origin: languages/vb/tests/vb/test_vb_exit_do.rs

Module M
    Function ReachedLimit(value As Integer) As Boolean
        Return value >= 5
    End Function

    Sub Main()
        Dim total As Integer = 0
        Do
            total = total + 2
            If ReachedLimit(total) Then Exit Do
        Loop
        Console.WriteLine(total)
    End Sub
End Module
