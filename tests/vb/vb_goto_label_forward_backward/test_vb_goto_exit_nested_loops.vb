' vybe-test: vb/vb_goto_label_forward_backward/test_vb_goto_exit_nested_loops
' origin: languages/vb/tests/vb/test_vb_goto_label_forward_backward.rs

Module Program
    Sub Main()
        For r As Integer = 1 To 5
            For c As Integer = 1 To 5
                If r = 2 AndAlso c = 2 Then GoTo BreakAll
            Next
        Next
BreakAll:
        Console.WriteLine("Broke Out of All Loops")
    End Sub
End Module
