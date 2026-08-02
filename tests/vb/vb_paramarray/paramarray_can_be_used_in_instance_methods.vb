' vybe-test: vb/vb_paramarray/paramarray_can_be_used_in_instance_methods
' origin: languages/vb/tests/vb/test_vb_paramarray.rs

Class TextWriter
    Public Function JoinWithBar(ParamArray values() As String) As String
        Dim result As String = ""
        For i As Integer = 0 To values.Length - 1
            If i > 0 Then
                result = result & "|"
            End If
            result = result & values(i)
        Next
        Return result
    End Function
End Class

Module M
    Sub Main()
        Dim writer As New TextWriter()
        Console.WriteLine(writer.JoinWithBar("a", "b", "c"))
    End Sub
End Module
