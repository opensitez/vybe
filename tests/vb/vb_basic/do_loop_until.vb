' vybe-test: vb/vb_basic/do_loop_until
' origin: languages/vb/tests/vb/vb_basic_test.rs

Module Program
    Sub Main()
        Dim i As Integer = 0
        Do
            Console.WriteLine(i)
            i = i + 1
        Loop Until i >= 3
    End Sub
End Module
