' vybe-test: vb/vb_linq_union_intersect_except_comparer/test_vb_linq_intersect_by_key_selector
' origin: languages/vb/tests/vb/test_vb_linq_union_intersect_except_comparer.rs

Imports System.Linq

Class Item
    Public Property ID As Integer
    Public Sub New(id As Integer) : Me.ID = id : End Sub
End Class

Module Program
    Sub Main()
        Dim list1 = {New Item(1), New Item(2), New Item(3)}
        Dim keys2 = {2, 3, 4}
        Dim res = list1.IntersectBy(keys2, Function(i) i.ID)
        For Each i In res
            Console.WriteLine(i.ID)
        Next
    End Sub
End Module
