' vybe-test: vb/vb_do_until_loop_while/do_until_loop_while
' origin: languages/vb/tests/vb/test_vb_do_until_loop_while.rs

Module M
    Sub Main()
        Dim i = 0
        
        ' Technically valid syntax in VB to mix conditions on Do and Loop
        Do Until i = 10
            i += 1
        Loop While i < 5
        
        Console.WriteLine(i)
    End Sub
End Module
