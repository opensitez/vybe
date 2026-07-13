use super::helpers::run_vb;

#[test]
fn anonymous_type_key() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Anonymous type with Key properties (makes them read-only and participates in Equals/GetHashCode)
        Dim a1 = New With {Key .Id = 1, .Name = "A"}
        Dim a2 = New With {Key .Id = 1, .Name = "B"}
        
        Console.WriteLine(a1.Id)
        Console.WriteLine(a1.Equals(a2)) ' Should be true because only Key properties are compared
        
        a1.Name = "C" ' Non-key is mutable
        Console.WriteLine(a1.Name)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["1", "True", "C"]);
}
