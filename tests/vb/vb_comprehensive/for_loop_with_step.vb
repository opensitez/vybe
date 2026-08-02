' vybe-test: vb/vb_comprehensive/for_loop_with_step
' origin: languages/vb/tests/vb/vb_comprehensive_test.rs

Module M
    Sub Main()
        For i As Integer = 0 To 10 Step 3
            Console.WriteLine(i)
        Next
    End Sub
End Module
