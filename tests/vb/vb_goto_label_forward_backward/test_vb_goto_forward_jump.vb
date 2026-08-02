' vybe-test: vb/vb_goto_label_forward_backward/test_vb_goto_forward_jump
' origin: languages/vb/tests/vb/test_vb_goto_label_forward_backward.rs

Module Program
    Sub Main()
        Console.WriteLine("Start")
        GoTo TargetLabel
        Console.WriteLine("Skipped")
TargetLabel:
        Console.WriteLine("End")
    End Sub
End Module
