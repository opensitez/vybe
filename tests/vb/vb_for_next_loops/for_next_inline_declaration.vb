' vybe-test: vb/vb_for_next_loops/for_next_inline_declaration
' origin: languages/vb/tests/vb/test_vb_for_next_loops.rs

Module M
Sub Main()
Dim sum = 0
For i As Integer = 1 To 3
sum += i
Next
Console.WriteLine(sum)
End Sub
End Module
