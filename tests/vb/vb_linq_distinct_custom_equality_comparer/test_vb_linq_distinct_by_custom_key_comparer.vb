' vybe-test: vb/vb_linq_distinct_custom_equality_comparer/test_vb_linq_distinct_by_custom_key_comparer
' origin: languages/vb/tests/vb/test_vb_linq_distinct_custom_equality_comparer.rs

Imports System
Imports System.Linq

Class Employee
    Public Property Department As String
    Public Property Name As String
    Public Sub New(d As String, n As String) : Department = d : Name = n : End Sub
End Class

Module Program
    Sub Main()
        Dim emps = {New Employee("hr", "Alice"), New Employee("HR", "Bob"), New Employee("IT", "Charlie")}
        Dim uniqueDepts = emps.DistinctBy(Function(e) e.Department, StringComparer.OrdinalIgnoreCase)
        For Each e In uniqueDepts
            Console.WriteLine(e.Department & ":" & e.Name)
        Next
    End Sub
End Module
