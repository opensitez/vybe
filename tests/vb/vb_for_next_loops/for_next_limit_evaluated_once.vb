' vybe-test: vb/vb_for_next_loops/for_next_limit_evaluated_once
' origin: languages/vb/tests/vb/test_vb_for_next_loops.rs

Module M
Dim limit As Integer = 3
Function GetLimit() As Integer
limit += 1
Return limit
End Function
Sub Main()
Dim sum = 0
For i = 1 To GetLimit()
sum += 1
Next
Console.WriteLine(sum)
End Sub
End Module
