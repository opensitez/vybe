use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: LINQ Group By Multi-Key & Group Into
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_linq_group_by_composite_key() {
    let src = r#"
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
"#;
    assert_eq!(
        run_vb(src),
        vec!["3", "IT-Dev:2", "IT-QA:1", "HR-Recruiter:1"]
    );
}

#[test]
fn test_vb_linq_group_by_aggregations() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim sales = {
            New With {.Region = "North", .Amount = 100D},
            New With {.Region = "North", .Amount = 200D},
            New With {.Region = "South", .Amount = 150D}
        }

        Dim summary = From s In sales
                      Group s By s.Region Into Total = Sum(s.Amount), Average = Average(s.Amount), Count()

        For Each sum In summary
            Console.WriteLine(sum.Region & ": Total=" & sum.Total & ", Count=" & sum.Count)
        Next
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["North: Total=300, Count=2", "South: Total=150, Count=1"]
    );
}
