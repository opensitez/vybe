' vybe-test: vb/vb_for_next_loops/for_next_floating_point
' origin: languages/vb/tests/vb/test_vb_for_next_loops.rs

Module M
Sub Main()
Dim sum As Double = 0
For i As Double = 0 To 1 Step 0.5
sum += i
Next
Console.WriteLine(sum)
End Sub
End Module
