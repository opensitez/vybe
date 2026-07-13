use super::helpers::run_vb;

#[test]
fn nameof_generic_type() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        ' NameOf with generic type
        Console.WriteLine(NameOf(List(Of Integer)))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["List"]);
}
