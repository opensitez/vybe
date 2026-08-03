use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: String Trim, TrimStart & TrimEnd Full Surface
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_string_trim_no_args_whitespace() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "   Hello World   "
        Console.WriteLine("'" & s.Trim() & "'")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["'Hello World'"]);
}

#[test]
fn test_vb_string_trim_tabs_and_newlines() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = vbTab & vbCrLf & "Data" & vbTab & vbCrLf
        Console.WriteLine("'" & s.Trim() & "'")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["'Data'"]);
}

#[test]
fn test_vb_string_trimstart_no_args() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "   LeftPadded   "
        Console.WriteLine("'" & s.TrimStart() & "'")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["'LeftPadded   '"]);
}

#[test]
fn test_vb_string_trimend_no_args() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "   RightPadded   "
        Console.WriteLine("'" & s.TrimEnd() & "'")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["'   RightPadded'"]);
}

#[test]
fn test_vb_string_trim_single_char() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "---Title---"
        Console.WriteLine(s.Trim("-"c))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Title"]);
}

#[test]
fn test_vb_string_trimstart_single_char() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "000012345"
        Console.WriteLine(s.TrimStart("0"c))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["12345"]);
}

#[test]
fn test_vb_string_trimend_single_char() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "123450000"
        Console.WriteLine(s.TrimEnd("0"c))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["12345"]);
}

#[test]
fn test_vb_string_trim_char_array_multiple() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "---***Hello***---"
        Dim trimChars As Char() = {"-"c, "*"c}
        Console.WriteLine(s.Trim(trimChars))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Hello"]);
}

#[test]
fn test_vb_string_trimstart_char_array_multiple() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "---***Hello***---"
        Dim trimChars As Char() = {"-"c, "*"c}
        Console.WriteLine(s.TrimStart(trimChars))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Hello***---"]);
}

#[test]
fn test_vb_string_trimend_char_array_multiple() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "---***Hello***---"
        Dim trimChars As Char() = {"-"c, "*"c}
        Console.WriteLine(s.TrimEnd(trimChars))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["---***Hello"]);
}

#[test]
fn test_vb_string_trim_non_matching_char() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "Hello"
        Console.WriteLine(s.Trim("x"c))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Hello"]);
}

#[test]
fn test_vb_string_trim_entire_string_matches() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "aaaaa"
        Console.WriteLine("'" & s.Trim("a"c) & "'")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["''"]);
}

#[test]
fn test_vb_string_trim_empty_string() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = ""
        Console.WriteLine("'" & s.Trim() & "'")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["''"]);
}

#[test]
fn test_vb_string_trim_punctuation_chars() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = ", ,Hello, World!..."
        Dim punct As Char() = {","c, "."c, "!"c}
        Console.WriteLine(s.Trim(punct))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Hello, World"]);
}

#[test]
fn test_vb_string_trim_unicode_whitespace() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = ChrW(&H2000) & "UnicodeSpace" & ChrW(&H2000)
        Console.WriteLine(s.Trim())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["UnicodeSpace"]);
}

#[test]
fn test_vb_string_trim_paramarray_syntax() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "<><>Data<><>"
        Console.WriteLine(s.Trim("<"c, ">"c))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Data"]);
}

#[test]
fn test_vb_string_trimstart_paramarray_syntax() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "<><>Data<><>"
        Console.WriteLine(s.TrimStart("<"c, ">"c))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Data<><>"]);
}

#[test]
fn test_vb_string_trimend_paramarray_syntax() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "<><>Data<><>"
        Console.WriteLine(s.TrimEnd("<"c, ">"c))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["<><>Data"]);
}

#[test]
fn test_vb_string_trim_case_sensitivity() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "aAaHelloaAa"
        Console.WriteLine(s.Trim("a"c))
        Console.WriteLine(s.Trim("A"c))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["AaHelloaA", "aAaHelloaAa"]);
}

#[test]
fn test_vb_string_trim_quotes() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = """QuotedValue"""
        Console.WriteLine(s.Trim(""""c))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["QuotedValue"]);
}
