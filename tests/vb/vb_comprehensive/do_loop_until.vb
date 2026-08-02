' vybe-test: vb/vb_comprehensive/do_loop_until
' origin: languages/vb/tests/vb/vb_comprehensive_test.rs

Module M
    Sub Main()
        Dim i As Integer = 0
        Do
            Console.WriteLine(i)
            i = i + 1
        Loop Until i >= 3
    End Sub
End Module
