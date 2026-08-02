' vybe-test: vb/vb_goto_label_forward_backward/test_vb_goto_out_of_do_while_loop
' origin: languages/vb/tests/vb/test_vb_goto_label_forward_backward.rs

Module Program
    Sub Main()
        Dim i = 0
        Do While True
            i += 1
            If i = 5 Then GoTo LoopExit
        Loop
LoopExit:
        Console.WriteLine("Exited Do Loop: " & i)
    End Sub
End Module
