' vybe-test: vb/vb_for_next_loops/for_next_decimal
' origin: languages/vb/tests/vb/test_vb_for_next_loops.rs

Module M
Sub Main()
Dim sum As Decimal = 0D
For i As Decimal = 0D To 1D Step 0.5D
sum += i
Next
Console.WriteLine(sum)
End Sub
End Module
