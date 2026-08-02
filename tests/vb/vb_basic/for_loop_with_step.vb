' vybe-test: vb/vb_basic/for_loop_with_step
' origin: languages/vb/tests/vb/vb_basic_test.rs

Module Program
    Sub Main()
        For i As Integer = 0 To 10 Step 2
            Console.WriteLine(i)
        Next
    End Sub
End Module
