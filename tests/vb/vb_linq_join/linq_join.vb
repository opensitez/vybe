' vybe-test: vb/vb_linq_join/linq_join
' origin: languages/vb/tests/vb/test_vb_linq_join.rs

Imports System.Linq

Class Employee
    Public Id As Integer
    Public Name As String
End Class

Class Dept
    Public Id As Integer
    Public DeptName As String
End Class

Module M
    Sub Main()
        Dim emps = {New Employee() With {.Id = 1, .Name = "Alice"}, New Employee() With {.Id = 2, .Name = "Bob"}}
        Dim depts = {New Dept() With {.Id = 1, .DeptName = "IT"}}
        
        Dim query = From e In emps
                    Join d In depts On e.Id Equals d.Id
                    Select e.Name, d.DeptName
                    
        For Each item In query
            Console.WriteLine(item.Name & "-" & item.DeptName)
        Next
    End Sub
End Module
