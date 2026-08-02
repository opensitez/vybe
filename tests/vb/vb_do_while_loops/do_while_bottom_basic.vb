' vybe-test: vb/vb_do_while_loops/do_while_bottom_basic
' origin: languages/vb/tests/vb/test_vb_do_while_loops.rs

Module M
Sub Main()
Dim x = 0
Do
x += 1
Loop While x < 3
Console.WriteLine(x)
End Sub
End Module
