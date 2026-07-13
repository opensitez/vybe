use super::helpers::run_vb;

#[test]
fn linq_let_complex() {
    let out = run_vb(
        r#"
Imports System.Linq

Module M
    Sub Main()
        Dim names = {"Alice", "Bob", "Charlie"}
        
        Dim query = From n In names
                    Let len = n.Length
                    Where len > 3
                    Select n, len
                    
        For Each item In query
            Console.WriteLine(item.n & item.len)
        Next
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Alice5", "Charlie7"]);
}
