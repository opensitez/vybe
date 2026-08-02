' vybe-test: vb/vb_do_while_loops/while_end_while_exit_while
' origin: languages/vb/tests/vb/test_vb_do_while_loops.rs

Module M
Sub Main()
Dim x = 0
While True
x += 1
If x = 3 Then Exit While
End While
Console.WriteLine(x)
End Sub
End Module
