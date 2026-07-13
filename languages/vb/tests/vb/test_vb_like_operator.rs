use super::helpers::run_vb;

#[test]
fn like_operator_wildcards() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim s1 As String = "Bat"
        Dim s2 As String = "Cat"
        Dim s3 As String = "Hat"
        
        ' ? matches any single character
        Console.WriteLine(s1 Like "?at")
        
        ' * matches zero or more characters
        Console.WriteLine(s2 Like "C*")
        
        ' # matches any single digit
        Console.WriteLine("123" Like "1#3")
        Console.WriteLine("1a3" Like "1#3")
        
        ' Character lists
        Console.WriteLine(s1 Like "[BCH]at")
        Console.WriteLine("Mat" Like "[BCH]at")
        
        ' Character list negation
        Console.WriteLine("Mat" Like "[!BCH]at")
    End Sub
End Module
"#,
    );
    assert_eq!(
        out,
        vec!["True", "True", "True", "False", "True", "False", "True"]
    );
}

#[test]
fn like_operator_ranges() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Character ranges
        Console.WriteLine("b" Like "[a-z]")
        Console.WriteLine("D" Like "[a-z]")
        Console.WriteLine("D" Like "[A-Z]")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["True", "False", "True"]);
}
