' vybe-test: vb/vb_do_while_loops/do_until_top_basic
' origin: languages/vb/tests/vb/test_vb_do_while_loops.rs

Module M
Sub Main()
Dim x = 0
Do Until x = 3
x += 1
Loop
Console.WriteLine(x)
End Sub
End Module
