use super::helpers::run_vb;

#[test]
fn strings_multiline_adv() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Multiline strings preserve whitespace
        Dim s As String = "Line 1
    Line 2
Line 3"
        Dim lines = s.Split({vbLf}, StringSplitOptions.None)
        Console.WriteLine(lines.Length)
        Console.WriteLine(lines(1).TrimStart())
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["3", "Line 2"]);
}

#[test]
fn strings_interpolation_multiline() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim name = "Alice"
        Dim s = $"Hello
{name}"
        Dim lines = s.Split({vbLf}, StringSplitOptions.None)
        Console.WriteLine(lines(0))
        Console.WriteLine(lines(1))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Hello", "Alice"]);
}
