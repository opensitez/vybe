' vybe-test: vb/vb_for_next_loops/for_next_variable_mutation
' origin: languages/vb/tests/vb/test_vb_for_next_loops.rs

Module M
Sub Main()
Dim sum = 0
For i = 1 To 5
sum += 1
i = 5
Next
Console.WriteLine(sum)
End Sub
End Module
