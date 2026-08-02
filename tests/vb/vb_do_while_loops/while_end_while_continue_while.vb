' vybe-test: vb/vb_do_while_loops/while_end_while_continue_while
' origin: languages/vb/tests/vb/test_vb_do_while_loops.rs

Module M
Sub Main()
Dim x = 0, sum = 0
While x < 4
x += 1
If x = 2 Then Continue While
sum += x
End While
Console.WriteLine(sum)
End Sub
End Module
