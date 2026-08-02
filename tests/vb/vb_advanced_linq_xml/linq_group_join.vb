' vybe-test: vb/vb_advanced_linq_xml/linq_group_join
' origin: languages/vb/tests/vb/test_vb_advanced_linq_xml.rs

Imports System.Linq
Imports System.Collections.Generic

Class Dept
    Public Id As Integer
    Public Name As String
End Class

Class Emp
    Public DeptId As Integer
    Public Name As String
End Class

Module M
    Sub Main()
        Dim depts = {New Dept With {.Id = 1, .Name = "IT"}}
        Dim emps = {New Emp With {.DeptId = 1, .Name = "Alice"}, New Emp With {.DeptId = 1, .Name = "Bob"}}
        
        Dim query = From d In depts
                    Group Join e In emps On d.Id Equals e.DeptId Into DeptEmps = Group
                    Select d.Name, Count = DeptEmps.Count()
                    
        For Each q In query
            Console.WriteLine(q.Name & "-" & q.Count)
        Next
    End Sub
End Module
