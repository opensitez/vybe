use super::helpers::run_vb;

#[test]
fn system_textbuilder_capacity_and_append_format() {
    let out = run_vb(
        r#"
Imports System.Text

Module M
    Sub Main()
        Dim sb As New StringBuilder(4)
        Console.WriteLine(sb.Capacity >= 4)
        Console.WriteLine(sb.Length)

        sb.AppendFormat("v={0}", 7)
        Console.WriteLine(sb.Length)
        Console.WriteLine(sb.ToString())
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "0", "3", "v=7"]);
}
