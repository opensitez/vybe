use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Nullable Types with LINQ
// ═══════════════════════════════════════════════════════════

#[test]
fn nullable_types_linq() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic
Imports System.Linq

Module M
    Sub Main()
        Dim nums As New List(Of Integer?) From { 1, Nothing, 3, 4, Nothing, 6 }
        
        ' Filter out nulls
        Dim query = From n In nums
                    Where n.HasValue
                    Select n.Value
                    
        For Each n In query
            Console.WriteLine(n)
        Next
        
        ' Sum with nulls (LINQ Sum handles nulls by ignoring them or throwing, depending on usage;
        ' in VB, calling Sum on IEnumerable(Of Integer?) returns Integer? and ignores nulls)
        Dim total = nums.Sum()
        Console.WriteLine("Total: " & total.ToString())
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["1", "3", "4", "6", "Total: 14"]);
}
