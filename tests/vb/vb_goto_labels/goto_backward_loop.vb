' vybe-test: vb/vb_goto_labels/goto_backward_loop
' origin: languages/vb/tests/vb/test_vb_goto_labels.rs

Module M
    Sub Main()
        Dim i As Integer = 0
        
LoopStart:
        If i = 3 Then
            GoTo Done
        End If
        
        Console.WriteLine(i)
        i = i + 1
        GoTo LoopStart
        
Done:
        Console.WriteLine("Finished")
    End Sub
End Module
