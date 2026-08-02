' vybe-test: vb/vb_goto_label_forward_backward/test_vb_goto_inside_if_statement_branch
' origin: languages/vb/tests/vb/test_vb_goto_label_forward_backward.rs

Module Program
    Sub Main()
        Dim flag = True
        If flag Then
            GoTo SuccessLabel
        Else
            GoTo FailureLabel
        End If
SuccessLabel:
        Console.WriteLine("Success Branch")
        Exit Sub
FailureLabel:
        Console.WriteLine("Failure Branch")
    End Sub
End Module
