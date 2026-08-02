' vybe-test: vb/vb_index_out_of_range_exception/test_vb_span_slice_bounds_check
' origin: languages/vb/tests/vb/test_vb_index_out_of_range_exception.rs

Imports System

Module Program
    Sub Main()
        Dim arr As Integer() = {10, 20, 30}
        Dim span As Span(Of Integer) = arr.AsSpan()
        Try
            Dim subSpan = span.Slice(1, 5)
            Console.WriteLine(subSpan.Length)
        Catch ex As ArgumentOutOfRangeException
            Console.WriteLine("Span.Slice ArgumentOutOfRangeException Caught")
        End Try
    End Sub
End Module
