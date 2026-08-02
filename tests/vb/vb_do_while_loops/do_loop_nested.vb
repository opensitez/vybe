' vybe-test: vb/vb_do_while_loops/do_loop_nested
' origin: languages/vb/tests/vb/test_vb_do_while_loops.rs

Module M
Sub Main()
Dim i = 0, count = 0
Do While i < 2
Dim j = 0
Do While j < 3
count += 1
j += 1
Loop
i += 1
Loop
Console.WriteLine(count)
End Sub
End Module
