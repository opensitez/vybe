' vybe-test: vb/vb_weak_reference_gc_collect/test_vb_weak_reference_array_of_weak_references
' origin: languages/vb/tests/vb/test_vb_weak_reference_gc_collect.rs

Imports System

Class Item
    Public Tag As String
    Public Sub New(t As String)
        Tag = t
    End Sub
End Class

Module Program
    Sub Main()
        Dim i1 As New Item("A")
        Dim i2 As New Item("B")
        Dim refs As WeakReference(Of Item)() = {New WeakReference(Of Item)(i1), New WeakReference(Of Item)(i2)}

        For Each r In refs
            Dim item As Item = Nothing
            r.TryGetTarget(item)
            Console.WriteLine(item.Tag)
        Next
    End Sub
End Module
