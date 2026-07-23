use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: LINQ Inner Join & Group Join (Left Outer Join)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_linq_inner_join_query_syntax() {
    let src = r#"
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
"#;
    assert_eq!(
        run_vb(src),
        vec![
            "Alice in Engineering",
            "Bob in Engineering",
            "Charlie in Marketing"
        ]
    );
}

#[test]
fn test_vb_linq_left_outer_join_group_join() {
    let src = r#"
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
"#;
    assert_eq!(run_vb(src), vec!["Engineering:Alice", "Sales:None"]);
}
