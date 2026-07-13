use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: LINQ (Let Keyword)
// ═══════════════════════════════════════════════════════════

#[test]
fn linq_let_clause() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim words As String() = {"apple", "banana", "cherry"}
        
        Dim query = From w In words
                    Let len = w.Length
                    Where len > 5
                    Select w & " is " & len.ToString()
                    
        For Each item In query
            Console.WriteLine(item)
        Next
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["banana is 6", "cherry is 6"]);
}
