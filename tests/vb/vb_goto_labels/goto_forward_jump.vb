' vybe-test: vb/vb_goto_labels/goto_forward_jump
' origin: languages/vb/tests/vb/test_vb_goto_labels.rs

Module M
    Sub Main()
        Console.WriteLine("Start")
        GoTo SkipThis
        
        Console.WriteLine("Should not print")
        
SkipThis:
        Console.WriteLine("End")
    End Sub
End Module
