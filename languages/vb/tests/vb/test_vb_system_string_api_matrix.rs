use super::helpers::run_vb;

#[test]
fn string_api_matrix_trim_and_case() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim s As String = "  VB-API  "
        Console.WriteLine(s.Trim())
        Console.WriteLine(s.Trim().Length)
        Console.WriteLine(s.ToUpperInvariant())
        Console.WriteLine(s.ToLowerInvariant())
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["VB-API", "6", "  VB-API  ", "  vb-api  "]);
}

#[test]
fn string_api_matrix_queries_and_indexing() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim text As String = "VBNET"
        Console.WriteLine(text.Length)
        Console.WriteLine(text.StartsWith("VB"))
        Console.WriteLine(text.EndsWith("NET"))
        Console.WriteLine(text.Contains("BN"))
        Console.WriteLine(text.IndexOf("N"))
        Console.WriteLine(text.LastIndexOf("N"))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["5", "True", "True", "True", "2", "2"]);
}

#[test]
fn string_api_matrix_replace_and_remove_slice() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim base As String = "a-b-c-d"
        Console.WriteLine(base.Replace("-", ""))
        Console.WriteLine(base.Substring(2, 3))
        Console.WriteLine(base.Remove(4))
        Console.WriteLine(base.Insert(0, "["))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["abcd", "b-c", "a-b-", "[a-b-c-d"]);
}

#[test]
fn string_api_matrix_split_join_roundtrip() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim text As String = "a,b;c,d"
        Dim pieces As String() = text.Split(","c, ";"c)

        Dim joined As String = String.Join("|", pieces)
        Console.WriteLine(pieces.Length)
        Console.WriteLine(joined)
        Console.WriteLine(pieces(2))
        Console.WriteLine(pieces(3))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["4", "a|b|c|d", "c", "d"]);
}

#[test]
fn string_api_matrix_formatting_contracts() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim a As String = String.Format("A={0}, B={1}", 1, "x")
        Dim b As String = String.Format("{0,5}", 42)
        Dim c As String = String.Format("{0:F1}", 3.14159)
        Console.WriteLine(a)
        Console.WriteLine(b)
        Console.WriteLine(c)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["A=1, B=x", "   42", "3.1"]);
}

#[test]
fn string_api_matrix_comparisons_and_empty_checks() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine(String.IsNullOrEmpty(""))
        Console.WriteLine(String.IsNullOrWhiteSpace("  "))
        Console.WriteLine(String.IsNullOrWhiteSpace("x"))
        Console.WriteLine(String.Equals("a", "A"))
        Console.WriteLine(String.Equals("a", "a"))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True", "False", "False", "True"]);
}

#[test]
fn string_api_matrix_copy_and_padding() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim left As String = "abc"
        Dim right As String = "def"
        Dim padded As String = left.PadLeft(5, "_")
        Dim chars() As Char = {"a"c, "b"c, "c"c}
        Dim copied As String = New String(chars)

        Console.WriteLine(left = copied)
        Console.WriteLine(left.PadRight(6, "-"))
        Console.WriteLine(padded)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "abc---", "__abc"]);
}
