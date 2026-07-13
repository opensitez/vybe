use super::helpers::run_vb;

#[test]
fn linq_aggregate() {
    let out = run_vb(
        r#"
Imports System.Linq

Module M
    Sub Main()
        Dim numbers = {1, 2, 3, 4}
        
        Dim query = Aggregate n In numbers Into Sum(), Max(), Min()
        
        Console.WriteLine(query.Sum)
        Console.WriteLine(query.Max)
        Console.WriteLine(query.Min)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["10", "4", "1"]);
}
