' vybe-test: vb/vb_control_flow/for_each_array
' origin: languages/vb/tests/vb/test_vb_control_flow.rs

Module M
    Sub Main()
        Dim arr() As Integer = {10, 20, 30}
        Dim sum As Integer = 0
        For Each x As Integer In arr
            sum = sum + x
        Next
        Console.WriteLine(sum)
    End Sub
End Module
