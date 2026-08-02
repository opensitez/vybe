' vybe-test: vb/vb_spec_object_model/object_model_spec_object_collection_property_can_be_iterated
' origin: languages/vb/tests/vb/test_vb_spec_object_model.rs

Class Bag
    Public Property Items As List(Of String)
End Class
Module M
    Sub Main()
        Dim bag As New Bag()
        bag.Items = New List(Of String)()
        bag.Items.Add("a")
        bag.Items.Add("b")
        Dim text As String = ""
        For Each item In bag.Items
            text &= item
        Next
        Console.WriteLine(text)
    End Sub
End Module
