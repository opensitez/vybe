' vybe-test: vb/vb_anonymous_type_generic_inference/test_vb_anonymous_type_array_of_interfaces_projection
' origin: languages/vb/tests/vb/test_vb_anonymous_type_generic_inference.rs

Imports System.Linq

Interface IIdentifiable
    ReadOnly Property ID As Integer
End Interface

Class Item
    Implements IIdentifiable
    Public ReadOnly Property ID As Integer Implements IIdentifiable.ID
    Public Sub New(id As Integer) : Me.ID = id : End Sub
End Class

Module Program
    Sub Main()
        Dim items As IIdentifiable() = {New Item(1), New Item(2)}
        Dim projected = items.Select(Function(i) New With {.ItemID = i.ID})
        For Each p In projected
            Console.WriteLine("ID:" & p.ItemID)
        Next
    End Sub
End Module
