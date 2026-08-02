' vybe-test: vb/vb_for_next_loops/for_next_dynamic_limits
' origin: languages/vb/tests/vb/test_vb_for_next_loops.rs

Module M
Function GetLimit() As Integer
Return 3
End Function
Sub Main()
Dim sum = 0
For i = 1 To GetLimit()
sum += 1
Next
Console.WriteLine(sum)
End Sub
End Module
