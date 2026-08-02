' vybe-test: vb/vb_linq_group_by_multi_key/test_vb_linq_group_by_composite_key
' origin: languages/vb/tests/vb/test_vb_linq_group_by_multi_key.rs

Imports System.Linq

Module Program
    Sub Main()
        Dim employees = {
            New With {.Dept = "IT", .Role = "Dev", .Name = "Alice"},
            New With {.Dept = "IT", .Role = "Dev", .Name = "Bob"},
            New With {.Dept = "IT", .Role = "QA", .Name = "Charlie"},
            New With {.Dept = "HR", .Role = "Recruiter", .Name = "David"}
        }

        Dim groups = From emp In employees
                     Group emp By Key = New With {emp.Dept, emp.Role} Into Group

        Console.WriteLine(groups.Count())
        For Each g In groups
            Console.WriteLine(g.Key.Dept & "-" & g.Key.Role & ":" & g.Group.Count())
        Next
    End Sub
End Module
