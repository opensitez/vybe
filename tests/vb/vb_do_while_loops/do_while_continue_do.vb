' vybe-test: vb/vb_do_while_loops/do_while_continue_do
' origin: languages/vb/tests/vb/test_vb_do_while_loops.rs

Module M
Sub Main()
Dim x = 0, sum = 0
Do While x < 4
x += 1
If x = 2 Then Continue Do
sum += x
Loop
Console.WriteLine(sum)
End Sub
End Module
