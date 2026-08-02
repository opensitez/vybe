' vybe-test: vb/vb_goto_label_forward_backward/test_vb_goto_label_with_same_name_as_variable
' origin: languages/vb/tests/vb/test_vb_goto_label_forward_backward.rs

Module Program
    Sub Main()
        Dim Target As String = "VarVal"
        GoTo Target
        Console.WriteLine("Skipped")
Target:
        Console.WriteLine(Target)
    End Sub
End Module
