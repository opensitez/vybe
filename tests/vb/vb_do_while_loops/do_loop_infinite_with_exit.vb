' vybe-test: vb/vb_do_while_loops/do_loop_infinite_with_exit
' origin: languages/vb/tests/vb/test_vb_do_while_loops.rs

Module M
Sub Main()
Dim x = 0
Do
x += 1
If x = 3 Then Exit Do
Loop
Console.WriteLine(x)
End Sub
End Module
