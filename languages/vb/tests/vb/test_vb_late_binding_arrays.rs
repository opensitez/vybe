use super::helpers::run_vb;

#[test]
fn late_binding_arrays() {
    let out = run_vb(
        r#"
Option Strict Off

Module M
    Sub Main()
        Dim obj As Object = New Integer() {1, 2, 3}
        
        ' Late bound array indexing
        Console.WriteLine(obj(1))
        
        obj(2) = 10
        Console.WriteLine(obj(2))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["2", "10"]);
}
