use super::helpers::run_vb;

#[test]
fn string_left_right_mid() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim s As String = "Hello World"
        
        ' 1-based indexing for these legacy string functions
        Console.WriteLine(Left(s, 5))
        Console.WriteLine(Right(s, 5))
        Console.WriteLine(Mid(s, 7, 5))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Hello", "World", "World"]);
}
