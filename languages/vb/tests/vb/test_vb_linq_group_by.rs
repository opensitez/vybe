use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: LINQ (Group By)
// ═══════════════════════════════════════════════════════════

#[test]
fn linq_group_by_clause() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim words As String() = {"apple", "ant", "banana", "bat", "cherry"}
        
        ' Group By generates a Key and a Group collection
        Dim query = From w In words
                    Group By firstLetter = w(0) Into Group
                    Select Key = firstLetter, Count = Group.Count()
                    
        For Each item In query
            Console.WriteLine(item.Key & ":" & item.Count.ToString())
        Next
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["a:2", "b:2", "c:1"]);
}
