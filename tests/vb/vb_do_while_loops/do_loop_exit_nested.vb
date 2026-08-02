' vybe-test: vb/vb_do_while_loops/do_loop_exit_nested
' origin: languages/vb/tests/vb/test_vb_do_while_loops.rs

Module M
Sub Main()
Dim i = 0, count = 0
Do While i < 3
Dim j = 0
Do While j < 3
j += 1
If j = 2 Then Exit Do
count += 1
Loop
i += 1
Loop
Console.WriteLine(count)
End Sub
End Module
