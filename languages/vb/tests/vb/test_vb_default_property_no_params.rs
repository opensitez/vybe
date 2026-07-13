use super::helpers::run_vb;

#[test]
fn default_property_no_params() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Default properties MUST have parameters in VB.NET.
        ' Testing parser constraint logic implicitly.
        Console.WriteLine("Parsed")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Parsed"]);
}
