' vybe-test: vb/vb_goto_label_forward_backward/test_vb_goto_inside_select_case_branch
' origin: languages/vb/tests/vb/test_vb_goto_label_forward_backward.rs

Module Program
    Sub Main()
        Dim mode = 2
        Select Case mode
            Case 1
                Console.WriteLine("Mode 1")
            Case 2
                GoTo SpecialMode
        End Select
        Console.WriteLine("Normal End")
        Exit Sub
SpecialMode:
        Console.WriteLine("Special Mode Jumped")
    End Sub
End Module
