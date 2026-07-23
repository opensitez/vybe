use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: ValueTuple Comparison & Sorting
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_tuple_equality_operators() {
    let src = r#"
Module Program
    Sub Main()
        Dim t1 = (1, "A")
        Dim t2 = (1, "A")
        Dim t3 = (1, "B")

        Console.WriteLine(t1 = t2)
        Console.WriteLine(t1 <> t3)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "True"]);
}

#[test]
fn test_vb_tuple_list_sort_lexicographical() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim list As New List(Of (Integer, String)) From {
            (2, "B"),
            (1, "Z"),
            (1, "A")
        }
        list.Sort()

        For Each item In list
            Console.WriteLine(item.Item1 & ":" & item.Item2)
        Next
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1:A", "1:Z", "2:B"]);
}
