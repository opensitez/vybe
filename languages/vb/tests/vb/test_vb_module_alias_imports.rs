use super::helpers::run_vb;

#[test]
fn module_alias_imports() {
    let out = run_vb(
        r#"
' Alias imports
Imports Txt = System.Text

Module M
    Sub Main()
        Dim sb As New Txt.StringBuilder()
        sb.Append("Alias")
        Console.WriteLine(sb.ToString())
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Alias"]);
}
