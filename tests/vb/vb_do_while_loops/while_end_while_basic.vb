' vybe-test: vb/vb_do_while_loops/while_end_while_basic
' origin: languages/vb/tests/vb/test_vb_do_while_loops.rs

Module M
Sub Main()
Dim x = 0
While x < 3
x += 1
End While
Console.WriteLine(x)
End Sub
End Module
