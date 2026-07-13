use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Array Methods (Filter)
// ═══════════════════════════════════════════════════════════

#[test]
fn array_filter_function() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim source As String() = {"apple", "banana", "apricot", "cherry"}
        
        ' Filter returns a new array with elements containing the match string
        Dim result As String() = Filter(source, "ap")
        
        For Each item In result
            Console.WriteLine(item)
        Next
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["apple", "apricot"]);
}

#[test]
fn array_filter_exclude() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim source As String() = {"apple", "banana", "apricot", "cherry"}
        
        ' Filter with Include=False returns elements that DO NOT contain the match
        Dim result As String() = Filter(source, "ap", False)
        
        For Each item In result
            Console.WriteLine(item)
        Next
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["banana", "cherry"]);
}
