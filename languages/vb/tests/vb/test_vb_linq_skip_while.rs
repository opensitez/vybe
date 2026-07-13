use super::helpers::run_vb;

#[test]
fn linq_skip_while() {
    let out = run_vb(
        r#"
Imports System.Linq

Module M
    Sub Main()
        Dim nums = {1, 2, 3, 4, 1, 2}
        
        Dim query = From n In nums
                    Skip While n < 3
                    Select n
                    
        For Each n In query
            Console.WriteLine(n)
        Next
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["3", "4", "1", "2"]);
}
