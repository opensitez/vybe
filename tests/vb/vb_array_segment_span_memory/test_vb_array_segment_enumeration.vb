' vybe-test: vb/vb_array_segment_span_memory/test_vb_array_segment_enumeration
' origin: languages/vb/tests/vb/test_vb_array_segment_span_memory.rs

Module Program
    Sub Main()
        Dim numbers As Integer() = {10, 20, 30, 40, 50}
        Dim segment As New ArraySegment(Of Integer)(numbers, 1, 3)
        Dim sum As Integer = 0
        For Each val In segment
            sum += val
        Next
        Console.WriteLine(sum)
    End Sub
End Module
