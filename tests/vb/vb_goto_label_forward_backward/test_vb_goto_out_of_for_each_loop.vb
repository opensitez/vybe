' vybe-test: vb/vb_goto_label_forward_backward/test_vb_goto_out_of_for_each_loop
' origin: languages/vb/tests/vb/test_vb_goto_label_forward_backward.rs

Module Program
    Sub Main()
        Dim items As String() = {"A", "B", "C"}
        For Each item In items
            If item = "B" Then GoTo FoundB
        Next
        Exit Sub
FoundB:
        Console.WriteLine("Found B via GoTo")
    End Sub
End Module
