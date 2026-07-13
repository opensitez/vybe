use super::helpers::run_vb;

#[test]
fn string_char_literals() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Double quotes is a String
        Dim s As String = "A"
        
        ' Double quotes followed by c is a Char
        Dim c As Char = "A"c
        
        Console.WriteLine(s.GetType().Name)
        Console.WriteLine(c.GetType().Name)
        
        ' Escaping double quotes inside a string literal
        Dim q As String = "He said ""Hello"" to me"
        Console.WriteLine(q)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["String", "Char", "He said \"Hello\" to me"]);
}
