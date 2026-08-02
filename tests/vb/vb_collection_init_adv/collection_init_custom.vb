' vybe-test: vb/vb_collection_init_adv/collection_init_custom
' origin: languages/vb/tests/vb/test_vb_collection_init_adv.rs

Imports System.Collections
Imports System.Collections.Generic

Class MyCol
    Implements IEnumerable(Of Integer)
    
    Private items As New List(Of Integer)
    
    Public Sub Add(val As Integer)
        items.Add(val * 2)
    End Sub
    
    Public Iterator Function GetEnumerator() As IEnumerator(Of Integer) Implements IEnumerable(Of Integer).GetEnumerator
        For Each item In items
            Yield item
        Next
    End Function

    Private Iterator Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        For Each item In items
            Yield item
        Next
    End Function
End Class

Module M
    Sub Main()
        Dim c As New MyCol From { 1, 2, 3 }
        Dim sum = 0
        For Each i In c
            sum += i
        Next
        Console.WriteLine(sum)
    End Sub
End Module
