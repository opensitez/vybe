' vybe-test: vb/vb_interop/b74_array_operations
' origin: languages/vb/tests/vb/vb_interop_test.rs

Dim arr(4) As Integer
For i As Integer = 0 To 4
    arr(i) = i * 10
Next
Console.WriteLine(arr(0))
Console.WriteLine(arr(2))
Console.WriteLine(arr(4))
