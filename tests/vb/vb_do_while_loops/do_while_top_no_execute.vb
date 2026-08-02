' vybe-test: vb/vb_do_while_loops/do_while_top_no_execute
' origin: languages/vb/tests/vb/test_vb_do_while_loops.rs

Module M
Sub Main()
Dim x = 5
Do While x < 3
x += 1
Loop
Console.WriteLine(x)
End Sub
End Module
