use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Null-Conditional Indexer
// ═══════════════════════════════════════════════════════════

#[test]
fn null_conditional_indexer() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim dict As Dictionary(Of String, String) = Nothing
        
        ' Null conditional array/indexer access
        ' It uses ?(index) in VB.NET (unlike ?[index] in C#)
        Dim val1 As String = dict?("Key")
        Console.WriteLine(val1 Is Nothing)
        
        dict = New Dictionary(Of String, String) From { {"Key", "Value"} }
        Dim val2 As String = dict?("Key")
        Console.WriteLine(val2)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["True", "Value"]);
}
