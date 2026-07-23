use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: LINQ ElementAt, ElementAtOrDefault, FirstOrDefault
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_linq_element_at_valid_and_out_of_bounds() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim arr = {10, 20, 30}
        Console.WriteLine(arr.ElementAt(1))
        Console.WriteLine(arr.ElementAtOrDefault(5))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["20", "0"]);
}

#[test]
fn test_vb_linq_first_or_default_custom_default_value() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim emptyList As New System.Collections.Generic.List(Of String)()
        Dim firstVal As String = emptyList.FirstOrDefault()
        Console.WriteLine(firstVal Is Nothing)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}
