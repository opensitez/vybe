' vybe-test: vb/vb_line_continuations/line_continuation_can_wrap_comparison_expression
' origin: languages/vb/tests/vb/test_vb_line_continuations.rs

Module M
    Sub Main()
        Dim result As Boolean = 1 + _
            2 = _
            3
        If result Then
            Console.WriteLine("match")
        Else
            Console.WriteLine("miss")
        End If
    End Sub
End Module
