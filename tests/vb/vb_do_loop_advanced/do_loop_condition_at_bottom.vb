' vybe-test: vb/vb_do_loop_advanced/do_loop_condition_at_bottom
' origin: languages/vb/tests/vb/test_vb_do_loop_advanced.rs

Module M
    Sub Main()
        Dim i = 10
        Do
            Console.WriteLine(i)
            i += 1
        Loop While i < 5 ' Executes at least once
        
        Dim j = 10
        Do
            Console.WriteLine(j)
            j += 1
        Loop Until j > 5 ' Evaluates true on first check, so stops
    End Sub
End Module
