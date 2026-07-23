use super::helpers::run_vb;

#[test]
fn system_regex_static_is_match_and_match_value() {
    let out = run_vb(
        r#"
Imports System.Text.RegularExpressions

Module M
    Sub Main()
        Console.WriteLine(Regex.IsMatch("invoice-42", "\d+"))
        Console.WriteLine(Regex.Match("item99", "\d+").Value)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "99"]);
}

#[test]
fn system_regex_static_replace() {
    let out = run_vb(
        r#"
Imports System.Text.RegularExpressions

Module M
    Sub Main()
        Dim result As String = Regex.Replace("a-b-c", "-", ":")
        Console.WriteLine(result)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["a:b:c"]);
}
