' vybe-test: vb/vb_do_loop_advanced/do_loop_while_until
' origin: languages/vb/tests/vb/test_vb_do_loop_advanced.rs

Module M
    Sub Main()
        Dim i = 0
        Do While i < 3
            Console.WriteLine(i)
            i += 1
        Loop
        
        Dim j = 0
        Do Until j = 2
            Console.WriteLine(j)
            j += 1
        Loop
    End Sub
End Module
