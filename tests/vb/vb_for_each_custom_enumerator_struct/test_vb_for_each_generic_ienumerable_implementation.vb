' vybe-test: vb/vb_for_each_custom_enumerator_struct/test_vb_for_each_generic_ienumerable_implementation
' origin: languages/vb/tests/vb/test_vb_for_each_custom_enumerator_struct.rs

Imports System.Collections
Imports System.Collections.Generic

Class NumberCollection
    Implements IEnumerable(Of Integer)

    Public Function GetEnumerator() As IEnumerator(Of Integer) Implements IEnumerable(Of Integer).GetEnumerator
        Return New List(Of Integer) From {1, 3, 5}.GetEnumerator()
    End Function

    Private Function GetEnumerator1() As IEnumerator Implements IEnumerable.GetEnumerator
        Return GetEnumerator()
    End Function
End Class

Module Program
    Sub Main()
        Dim col As New NumberCollection()
        Dim total = 0
        For Each n In col
            total += n
        Next
        Console.WriteLine(total)
    End Sub
End Module
