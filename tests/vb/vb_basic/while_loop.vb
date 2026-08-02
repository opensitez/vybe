' vybe-test: vb/vb_basic/while_loop
' origin: languages/vb/tests/vb/vb_basic_test.rs

Module Program
    Sub Main()
        Dim i As Integer = 0
        While i < 5
            Console.WriteLine(i)
            i = i + 1
        End While
    End Sub
End Module
