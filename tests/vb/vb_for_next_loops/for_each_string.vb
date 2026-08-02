' vybe-test: vb/vb_for_next_loops/for_each_string
' origin: languages/vb/tests/vb/test_vb_for_next_loops.rs

Module M
Sub Main()
Dim s = "ABC"
Dim count = 0
For Each c In s
count += 1
Next
Console.WriteLine(count)
End Sub
End Module
