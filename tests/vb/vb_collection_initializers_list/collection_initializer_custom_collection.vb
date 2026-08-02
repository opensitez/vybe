' vybe-test: vb/vb_collection_initializers_list/collection_initializer_custom_collection
' origin: languages/vb/tests/vb/test_vb_collection_initializers_list.rs

Imports System.Collections.Generic

Class MyBag
    Implements IEnumerable(Of String)
    
    Private _items As New List(Of String)()
    
    ' Required Add method for collection initializer
    Public Sub Add(item As String)
        _items.Add("My" & item)
    End Sub
    
    Public Function GetEnumerator() As IEnumerator(Of String) Implements IEnumerable(Of String).GetEnumerator
        Return _items.GetEnumerator()
    End Function
    
    Private Function IEnumerable_GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
        Return _items.GetEnumerator()
    End Function
End Class

Module M
    Sub Main()
        Dim bag As New MyBag From {"Cat", "Dog"}
        
        For Each b In bag
            Console.WriteLine(b)
        Next
    End Sub
End Module
