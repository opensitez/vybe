use super::helpers::run_vb;

#[test]
fn query_let_where() {
    let out = run_vb(
        r#"
Imports System.Linq

Module M
    Sub Main()
        Dim numbers = {1, 2, 3, 4, 5}
        
        Dim query = From n In numbers
                    Let doubled = n * 2
                    Where doubled > 5
                    Select doubled
                    
        For Each d In query
            Console.WriteLine(d)
        Next
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["6", "8", "10"]);
}
