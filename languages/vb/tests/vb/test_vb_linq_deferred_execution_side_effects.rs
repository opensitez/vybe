use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: LINQ Deferred Execution & Evaluation Order
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_linq_deferred_execution_list_mutation() {
    let src = r#"
Imports System.Collections.Generic
Imports System.Linq

Module Program
    Sub Main()
        Dim numbers As New List(Of Integer) From {1, 2, 3}
        Dim query = From n In numbers Where n > 1 Select n * 10

        numbers.Add(4) ' Mutate source before enumeration

        Console.WriteLine(String.Join(",", query))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["20,30,40"]);
}

#[test]
fn test_vb_linq_immediate_execution_to_list() {
    let src = r#"
Imports System.Collections.Generic
Imports System.Linq

Module Program
    Sub Main()
        Dim numbers As New List(Of Integer) From {1, 2, 3}
        Dim snapshot = (From n In numbers Where n > 1 Select n * 10).ToList()

        numbers.Add(4) ' Source mutation does NOT affect materialized list

        Console.WriteLine(String.Join(",", snapshot))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["20,30"]);
}
