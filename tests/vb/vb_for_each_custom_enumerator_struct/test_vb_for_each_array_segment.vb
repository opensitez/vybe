' vybe-test: vb/vb_for_each_custom_enumerator_struct/test_vb_for_each_array_segment
' origin: languages/vb/tests/vb/test_vb_for_each_custom_enumerator_struct.rs

Imports System

Module Program
    Sub Main()
        Dim raw As Integer() = {10, 20, 30, 40, 50}
        Dim seg As New ArraySegment(Of Integer)(raw, 1, 3)
        Dim sum = 0
        For Each val In seg
            sum += val
        Next
        Console.WriteLine(sum)
    End Sub
End Module
