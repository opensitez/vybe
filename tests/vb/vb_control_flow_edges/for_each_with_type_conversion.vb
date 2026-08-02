' vybe-test: vb/vb_control_flow_edges/for_each_with_type_conversion
' origin: languages/vb/tests/vb/test_vb_control_flow_edges.rs

Module M
    Sub Main()
        Dim arr As Integer() = {1, 2, 3}
        ' For Each with implicit conversion to Double
        For Each x As Double In arr
            Console.WriteLine(x + 0.5)
        Next
    End Sub
End Module
