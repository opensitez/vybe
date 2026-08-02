' vybe-test: vb/vb_for_each_custom_enumerator_struct/test_vb_for_each_custom_duck_typed_collection
' origin: languages/vb/tests/vb/test_vb_for_each_custom_enumerator_struct.rs

Class DuckCollection
    Public Function GetEnumerator() As DuckEnumerator
        Return New DuckEnumerator()
    End Function
End Class

Class DuckEnumerator
    Private count As Integer = 0
    Public Function MoveNext() As Boolean
        count += 1
        Return count <= 2
    End Function
    Public ReadOnly Property Current As String
        Get
            Return "Quack" & count
        End Get
    End Property
End Class

Module Program
    Sub Main()
        Dim duckCol As New DuckCollection()
        For Each item In duckCol
            Console.WriteLine(item)
        Next
    End Sub
End Module
