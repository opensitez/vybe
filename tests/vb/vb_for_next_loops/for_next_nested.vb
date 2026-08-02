' vybe-test: vb/vb_for_next_loops/for_next_nested
' origin: languages/vb/tests/vb/test_vb_for_next_loops.rs

Module M
Sub Main()
Dim count = 0
For i = 1 To 2
For j = 1 To 3
count += 1
Next
Next
Console.WriteLine(count)
End Sub
End Module
