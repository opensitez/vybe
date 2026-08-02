' vybe-test: vb/vb_for_next_loops/for_each_collection
' origin: languages/vb/tests/vb/test_vb_for_next_loops.rs

Module M
Sub Main()
Dim col As New Collection()
col.Add(10)
col.Add(20)
Dim sum = 0
For Each v In col
sum += CInt(v)
Next
Console.WriteLine(sum)
End Sub
End Module
