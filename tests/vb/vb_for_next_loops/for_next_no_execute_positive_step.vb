' vybe-test: vb/vb_for_next_loops/for_next_no_execute_positive_step
' origin: languages/vb/tests/vb/test_vb_for_next_loops.rs

Module M
Sub Main()
Dim sum = 0
For i = 5 To 1
sum += i
Next
Console.WriteLine(sum)
End Sub
End Module
