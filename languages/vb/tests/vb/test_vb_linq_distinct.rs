use super::helpers::run_vb;

#[test]
fn linq_distinct() {
    let out = run_vb(
        r#"
Imports System.Linq

Module M
    Sub Main()
        Dim numbers = {1, 2, 2, 3, 3, 3}
        
        Dim query = (From n In numbers Select n).Distinct()
        
        For Each n In query
            Console.WriteLine(n)
        Next
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["1", "2", "3"]);
}
