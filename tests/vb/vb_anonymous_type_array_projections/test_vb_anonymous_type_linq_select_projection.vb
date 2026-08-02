' vybe-test: vb/vb_anonymous_type_array_projections/test_vb_anonymous_type_linq_select_projection
' origin: languages/vb/tests/vb/test_vb_anonymous_type_array_projections.rs

Imports System.Linq

Class Employee
    Public Property Name As String
    Public Property Salary As Double
    Public Sub New(n As String, s As Double) : Name = n : Salary = s : End Sub
End Class

Module Program
    Sub Main()
        Dim emps = {New Employee("Alice", 50000), New Employee("Bob", 60000)}
        Dim projected = From e In emps Select New With {.EmpName = e.Name, .AnnualSalary = e.Salary}
        For Each item In projected
            Console.WriteLine(item.EmpName & "=" & item.AnnualSalary)
        Next
    End Sub
End Module
