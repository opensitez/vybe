' vybe-test: vb/vb_anonymous_type_array_projections/test_vb_anonymous_type_linq_group_by_projection
' origin: languages/vb/tests/vb/test_vb_anonymous_type_array_projections.rs

Imports System.Linq

Class Student
    Public Property Grade As Integer
    Public Property Name As String
    Public Sub New(g As Integer, n As String) : Grade = g : Name = n : End Sub
End Class

Module Program
    Sub Main()
        Dim students = {New Student(10, "Alice"), New Student(10, "Bob"), New Student(11, "Charlie")}
        Dim groups = From s In students
                     Group s By s.Grade Into Group
                     Select New With {.Grade = Grade, .Count = Group.Count()}

        For Each g In groups
            Console.WriteLine("Grade " & g.Grade & " Count: " & g.Count)
        Next
    End Sub
End Module
