' vybe-test: vb/vb_goto_label_forward_backward/test_vb_goto_numeric_line_labels_legacy
' origin: languages/vb/tests/vb/test_vb_goto_label_forward_backward.rs

Module Program
    Sub Main()
10:     Console.WriteLine("Line 10")
        GoTo 30
20:     Console.WriteLine("Line 20")
30:     Console.WriteLine("Line 30")
    End Sub
End Module
