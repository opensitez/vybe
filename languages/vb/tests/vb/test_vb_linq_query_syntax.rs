use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: LINQ Query Syntax
// ═══════════════════════════════════════════════════════════

#[test]
fn linq_from_where_select() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim numbers As Integer() = {1, 2, 3, 4, 5, 6, 7, 8, 9, 10}
        
        Dim evens = From n In numbers
                    Where n Mod 2 = 0
                    Select n
                    
        For Each e In evens
            Console.WriteLine(e)
        Next
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["2", "4", "6", "8", "10"]);
}

#[test]
fn linq_order_by() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim words As String() = {"apple", "cherry", "banana"}
        
        Dim sorted = From w In words
                     Order By w Descending
                     Select w
                     
        For Each w In sorted
            Console.WriteLine(w)
        Next
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["cherry", "banana", "apple"]);
}

#[test]
fn linq_aggregate() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim numbers As Integer() = {1, 2, 3, 4, 5}
        
        ' Aggregate is a distinct keyword in VB LINQ
        Dim sum = Aggregate n In numbers Into Sum()
        Console.WriteLine(sum)
        
        Dim max = Aggregate n In numbers Into Max()
        Console.WriteLine(max)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["15", "5"]);
}
