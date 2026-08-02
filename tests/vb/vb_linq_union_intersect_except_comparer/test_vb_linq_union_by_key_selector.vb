' vybe-test: vb/vb_linq_union_intersect_except_comparer/test_vb_linq_union_by_key_selector
' origin: languages/vb/tests/vb/test_vb_linq_union_intersect_except_comparer.rs

Imports System.Linq

Class Person
    Public Property Name As String
    Public Sub New(n As String) : Name = n : End Sub
End Class

Module Program
    Sub Main()
        Dim list1 = {New Person("Alice"), New Person("Bob")}
        Dim list2 = {New Person("Bob"), New Person("Charlie")}
        Dim res = list1.UnionBy(list2, Function(p) p.Name)
        For Each p In res
            Console.WriteLine(p.Name)
        Next
    End Sub
End Module
