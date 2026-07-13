use super::helpers::run_vb;

#[test]
fn option_compare_text() {
    let out = run_vb(
        r#"
Option Compare Text

Module M
    Sub Main()
        Dim s1 As String = "HELLO"
        Dim s2 As String = "hello"
        
        ' With Option Compare Text, case-insensitive string comparison is used for =, <>, <, >, Like, etc.
        Console.WriteLine(s1 = s2)
        Console.WriteLine(s1 Like s2)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn option_compare_binary() {
    let out = run_vb(
        r#"
Option Compare Binary

Module M
    Sub Main()
        Dim s1 As String = "HELLO"
        Dim s2 As String = "hello"
        
        ' With Option Compare Binary, case-sensitive string comparison is used
        Console.WriteLine(s1 = s2)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["False"]);
}
