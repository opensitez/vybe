use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Advanced string operations
// ═══════════════════════════════════════════════════════════

#[test]
fn string_concat_ampersand() {
    let out = run_vb(r#"
Module M
    Sub Main()
        Dim a As String = "Hello"
        Dim b As String = "World"
        Console.WriteLine(a & " " & b)
    End Sub
End Module
"#);
    assert_eq!(out, vec!["Hello World"]);
}

#[test]
fn string_length() {
    let out = run_vb(r#"
Module M
    Sub Main()
        Console.WriteLine(Len("Hello"))
    End Sub
End Module
"#);
    assert_eq!(out, vec!["5"]);
}

#[test]
fn string_left_right_mid() {
    let out = run_vb(r#"
Module M
    Sub Main()
        Dim s As String = "Hello World"
        Console.WriteLine(Left(s, 5))
        Console.WriteLine(Right(s, 5))
        Console.WriteLine(Mid(s, 7, 5))
    End Sub
End Module
"#);
    assert_eq!(out, vec!["Hello", "World", "World"]);
}

#[test]
fn string_ucase_lcase() {
    let out = run_vb(r#"
Module M
    Sub Main()
        Console.WriteLine(UCase("hello"))
        Console.WriteLine(LCase("HELLO"))
    End Sub
End Module
"#);
    assert_eq!(out, vec!["HELLO", "hello"]);
}

#[test]
fn string_trim() {
    let out = run_vb(r#"
Module M
    Sub Main()
        Console.WriteLine(Trim("  hello  "))
        Console.WriteLine(LTrim("  hello"))
        Console.WriteLine(RTrim("hello  "))
    End Sub
End Module
"#);
    assert_eq!(out, vec!["hello", "hello", "hello"]);
}

#[test]
fn string_instr() {
    let out = run_vb(r#"
Module M
    Sub Main()
        Console.WriteLine(InStr("Hello World", "World"))
        Console.WriteLine(InStr("Hello", "xyz"))
    End Sub
End Module
"#);
    assert_eq!(out, vec!["7", "0"]);
}

#[test]
fn string_replace() {
    let out = run_vb(r#"
Module M
    Sub Main()
        Console.WriteLine(Replace("Hello World", "World", "VB"))
    End Sub
End Module
"#);
    assert_eq!(out, vec!["Hello VB"]);
}

#[test]
fn string_split_join() {
    let out = run_vb(r#"
Module M
    Sub Main()
        Dim parts() As String = Split("a,b,c", ",")
        Console.WriteLine(UBound(parts))
        Console.WriteLine(Join(parts, "-"))
    End Sub
End Module
"#);
    assert_eq!(out, vec!["2", "a-b-c"]);
}

#[test]
fn string_space_function() {
    let out = run_vb(r#"
Module M
    Sub Main()
        Console.WriteLine("[" & Space(5) & "]")
    End Sub
End Module
"#);
    assert_eq!(out, vec!["[     ]"]);
}

#[test]
fn string_chr_asc() {
    let out = run_vb(r#"
Module M
    Sub Main()
        Console.WriteLine(Asc("A"))
        Console.WriteLine(Chr(66))
    End Sub
End Module
"#);
    assert_eq!(out, vec!["65", "B"]);
}

#[test]
fn string_isnumeric() {
    let out = run_vb(r#"
Module M
    Sub Main()
        Console.WriteLine(IsNumeric("123"))
        Console.WriteLine(IsNumeric("abc"))
    End Sub
End Module
"#);
    assert_eq!(out, vec!["true", "false"]);
}

#[test]
fn string_cstr_val() {
    let out = run_vb(r#"
Module M
    Sub Main()
        Dim x As Integer = 42
        Console.WriteLine(CStr(x))
        Console.WriteLine(Val("123.45"))
    End Sub
End Module
"#);
    assert_eq!(out, vec!["42", "123.45"]);
}

#[test]
fn string_concat_with_number() {
    let out = run_vb(r#"
Module M
    Sub Main()
        Dim count As Integer = 5
        Console.WriteLine("Count: " & CStr(count))
    End Sub
End Module
"#);
    assert_eq!(out, vec!["Count: 5"]);
}

#[test]
fn string_comparison() {
    let out = run_vb(r#"
Module M
    Sub Main()
        If "abc" = "abc" Then Console.WriteLine("equal")
        If "abc" <> "xyz" Then Console.WriteLine("not equal")
    End Sub
End Module
"#);
    assert_eq!(out, vec!["equal", "not equal"]);
}

#[test]
fn string_multiline_concat() {
    let out = run_vb(r#"
Module M
    Sub Main()
        Dim s As String = "Hello" & _
            " " & _
            "World"
        Console.WriteLine(s)
    End Sub
End Module
"#);
    assert_eq!(out, vec!["Hello World"]);
}
