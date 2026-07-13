use super::helpers::run_vb;

#[test]
fn imports_primitive_alias() {
    let out = run_vb(
        r#"
Imports MyInt = System.Int32
Imports MyStr = System.String

Module M
    Sub Main()
        Dim i As MyInt = 42
        Dim s As MyStr = "Alias"
        Console.WriteLine(i)
        Console.WriteLine(s)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["42", "Alias"]);
}
