' vybe-test: vb/vb_paramarray/paramarray_preserves_string_order_when_joining
' origin: languages/vb/tests/vb/test_vb_paramarray.rs

Module M
    Function JoinAll(ParamArray values() As String) As String
        Dim result As String = ""
        For i As Integer = 0 To values.Length - 1
            If i > 0 Then
                result = result & ","
            End If
            result = result & values(i)
        Next
        Return result
    End Function

    Sub Main()
        Console.WriteLine(JoinAll("red", "green", "blue"))
    End Sub
End Module
