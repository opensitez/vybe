use super::helpers::run_vb;

#[test]
fn system_regex_match() {
    let out = run_vb(
        r#"
Imports System.Text.RegularExpressions

Module M
    Sub Main()
        Dim input As String = "The quick brown fox jumps over 42 lazy dogs."
        Dim pattern As String = "\b\d+\b"
        
        Dim m As Match = Regex.Match(input, pattern)
        If m.Success Then
            Console.WriteLine(m.Value)
        End If
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn system_regex_replace() {
    let out = run_vb(
        r#"
Imports System.Text.RegularExpressions

Module M
    Sub Main()
        Dim input As String = "apple banana apple"
        Dim pattern As String = "apple"
        
        Dim result As String = Regex.Replace(input, pattern, "orange")
        Console.WriteLine(result)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["orange banana orange"]);
}

#[test]
fn system_regex_split() {
    let out = run_vb(
        r#"
Imports System.Text.RegularExpressions

Module M
    Sub Main()
        Dim input As String = "a, b; c |d"
        Dim pattern As String = "[,;\|]\s*"
        
        Dim parts As String() = Regex.Split(input, pattern)
        Console.WriteLine(parts.Length)
        Console.WriteLine(parts(0))
        Console.WriteLine(parts(3))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["4", "a", "d"]);
}
