use super::helpers::run_vb;

#[test]
fn linq_group_by_into() {
    let out = run_vb(
        r#"
Imports System.Linq

Module M
    Sub Main()
        Dim words = {"apple", "banana", "apricot", "cherry"}
        
        Dim query = From w In words
                    Group By Key = w.Substring(0, 1) Into Group, Count()
                    Order By Key
                    
        For Each g In query
            Console.WriteLine(g.Key & g.Count)
        Next
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["a2", "b1", "c1"]);
}
