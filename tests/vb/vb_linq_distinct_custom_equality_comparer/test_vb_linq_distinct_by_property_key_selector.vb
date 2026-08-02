' vybe-test: vb/vb_linq_distinct_custom_equality_comparer/test_vb_linq_distinct_by_property_key_selector
' origin: languages/vb/tests/vb/test_vb_linq_distinct_custom_equality_comparer.rs

Imports System.Linq

Class Person
    Public Property Name As String
    Public Property Age As Integer
    Public Sub New(n As String, a As Integer) : Name = n : Age = a : End Sub
End Class

Module Program
    Sub Main()
        Dim people = {New Person("Alice", 25), New Person("Bob", 25), New Person("Charlie", 30)}
        Dim uniqueByAge = people.DistinctBy(Function(p) p.Age)
        For Each p In uniqueByAge
            Console.WriteLine(p.Name & ":" & p.Age)
        Next
    End Sub
End Module
