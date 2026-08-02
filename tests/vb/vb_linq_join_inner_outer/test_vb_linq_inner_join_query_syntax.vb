' vybe-test: vb/vb_linq_join_inner_outer/test_vb_linq_inner_join_query_syntax
' origin: languages/vb/tests/vb/test_vb_linq_join_inner_outer.rs

Imports System.Linq

Module Program
    Sub Main()
        Dim departments = {
            New With {.Id = 1, .Name = "Engineering"},
            New With {.Id = 2, .Name = "Marketing"}
        }

        Dim employees = {
            New With {.Name = "Alice", .DeptId = 1},
            New With {.Name = "Bob", .DeptId = 1},
            New With {.Name = "Charlie", .DeptId = 2}
        }

        Dim query = From emp In employees
                    Join dept In departments On emp.DeptId Equals dept.Id
                    Select emp.Name, dept.Name

        For Each item In query
            Console.WriteLine(item.Name & " in " & item.dept_Name)
        Next
    End Sub
End Module
