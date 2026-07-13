use super::helpers::run_vb;

#[test]
fn string_manipulation_methods() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim s As String = "  Hello, World!  "
        Console.WriteLine(s.Trim())
        Console.WriteLine(s.ToUpper().Trim())
        Console.WriteLine(s.ToLower().Trim())
        Console.WriteLine(s.Replace("World", "VB").Trim())
        Console.WriteLine(s.Substring(2, 5))
    End Sub
End Module
"#,
    );
    assert_eq!(
        out,
        vec![
            "Hello, World!",
            "HELLO, WORLD!",
            "hello, world!",
            "Hello, VB!",
            "Hello"
        ]
    );
}

#[test]
fn string_builder() {
    let out = run_vb(
        r#"
Imports System.Text

Module M
    Sub Main()
        Dim sb As New StringBuilder()
        sb.Append("Hello")
        sb.Append(" ")
        sb.Append("World")
        Console.WriteLine(sb.ToString())
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Hello World"]);
}
