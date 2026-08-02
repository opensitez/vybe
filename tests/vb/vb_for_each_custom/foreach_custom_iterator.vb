' vybe-test: vb/vb_for_each_custom/foreach_custom_iterator
' origin: languages/vb/tests/vb/test_vb_for_each_custom.rs

Imports System.Collections
Imports System.Collections.Generic

Class CustomCollection
    Implements IEnumerable(Of String)

    Private items As String() = {"Apple", "Banana", "Cherry"}

    Public Iterator Function GetEnumerator() As IEnumerator(Of String) Implements IEnumerable(Of String).GetEnumerator
        For i As Integer = 0 To items.Length - 1
            Yield items(i)
        Next
    End Function

    Private Iterator Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        For i As Integer = 0 To items.Length - 1
            Yield items(i)
        Next
    End Function
End Class

Module M
    Sub Main()
        Dim c As New CustomCollection()
        For Each item As String In c
            Console.WriteLine(item)
        Next
    End Sub
End Module
