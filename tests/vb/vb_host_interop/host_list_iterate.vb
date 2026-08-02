' vybe-test: vb/vb_host_interop/host_list_iterate
' origin: languages/vb/tests/vb/vb_host_interop_test.rs

Dim list As New List(Of String)
list.Add("x")
list.Add("y")
Dim total As Integer = 0
For Each item In list
    total = total + 1
Next
Console.WriteLine(total)
