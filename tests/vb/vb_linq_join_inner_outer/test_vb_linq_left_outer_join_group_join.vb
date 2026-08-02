' vybe-test: vb/vb_linq_join_inner_outer/test_vb_linq_left_outer_join_group_join
' origin: languages/vb/tests/vb/test_vb_linq_join_inner_outer.rs

Imports System.Linq

Module Program
    Sub Main()
        Dim departments = {
            New With {.Id = 1, .Name = "Engineering"},
            New With {.Id = 2, .Name = "Sales"}
        }

        Dim employees = {
            New With {.Name = "Alice", .DeptId = 1}
        }

        Dim query = From dept In departments
                    Group Join emp In employees On dept.Id Equals emp.DeptId Into Emps = Group
                    From emp In Emps.DefaultIfEmpty()
                    Select DeptName = dept.Name, EmpName = If(emp IsNot Nothing, emp.Name, "None")

        For Each item In query
            Console.WriteLine(item.DeptName & ":" & item.EmpName)
        Next
    End Sub
End Module
