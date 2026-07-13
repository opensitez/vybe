use super::helpers::run_vb;

#[test]
fn query_select_anonymous() {
    let out = run_vb(
        r#"
Imports System.Linq

Module M
    Sub Main()
        Dim ids = {1, 2}
        
        Dim query = From id In ids
                    Select New With {.Index = id, .Name = "Item" & id}
                    
        For Each item In query
            Console.WriteLine(item.Index & "-" & item.Name)
        Next
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["1-Item1", "2-Item2"]);
}
