use super::helpers::run_vb;

// ============================================================
// String builtins
// ============================================================

#[test]
fn builtin_left() {
    let out = run_vb(r#"
Module Program
    Sub Main()
        Console.WriteLine(Left("Hello World", 5))
    End Sub
End Module
"#);
    assert_eq!(out, vec!["Hello"]);
}

#[test]
fn builtin_right() {
    let out = run_vb(r#"
Module Program
    Sub Main()
        Console.WriteLine(Right("Hello World", 5))
    End Sub
End Module
"#);
    assert_eq!(out, vec!["World"]);
}

#[test]
fn builtin_mid() {
    let out = run_vb(r#"
Module Program
    Sub Main()
        Console.WriteLine(Mid("Hello World", 7))
        Console.WriteLine(Mid("Hello World", 7, 3))
    End Sub
End Module
"#);
    assert_eq!(out, vec!["World", "Wor"]);
}

#[test]
fn builtin_instr() {
    let out = run_vb(r#"
Module Program
    Sub Main()
        Console.WriteLine(InStr("Hello World", "World"))
        Console.WriteLine(InStr("Hello World", "xyz"))
    End Sub
End Module
"#);
    assert_eq!(out, vec!["7", "0"]);
}

#[test]
fn builtin_replace() {
    let out = run_vb(r#"
Module Program
    Sub Main()
        Console.WriteLine(Replace("Hello World", "World", "VB"))
    End Sub
End Module
"#);
    assert_eq!(out, vec!["Hello VB"]);
}

#[test]
fn builtin_split_join() {
    let out = run_vb(r#"
Module Program
    Sub Main()
        Dim parts() As String = Split("a,b,c", ",")
        Console.WriteLine(Join(parts, "-"))
    End Sub
End Module
"#);
    assert_eq!(out, vec!["a-b-c"]);
}

#[test]
fn builtin_ltrim_rtrim() {
    let out = run_vb(r#"
Module Program
    Sub Main()
        Console.WriteLine(LTrim("  hello"))
        Console.WriteLine(RTrim("hello  "))
    End Sub
End Module
"#);
    assert_eq!(out, vec!["hello", "hello"]);
}

#[test]
fn builtin_asc_chr() {
    let out = run_vb(r#"
Module Program
    Sub Main()
        Console.WriteLine(Asc("A"))
        Console.WriteLine(Chr(65))
    End Sub
End Module
"#);
    assert_eq!(out, vec!["65", "A"]);
}

#[test]
fn builtin_space() {
    let out = run_vb(r#"
Module Program
    Sub Main()
        Console.WriteLine(">" & Space(3) & "<")
    End Sub
End Module
"#);
    assert_eq!(out, vec![">   <"]);
}

// ============================================================
// Conversion builtins
// ============================================================

#[test]
fn builtin_cstr() {
    let out = run_vb(r#"
Module Program
    Sub Main()
        Console.WriteLine(CStr(42))
    End Sub
End Module
"#);
    assert_eq!(out, vec!["42"]);
}

#[test]
fn builtin_cint() {
    let out = run_vb(r#"
Module Program
    Sub Main()
        Console.WriteLine(CInt(3.7))
    End Sub
End Module
"#);
    assert_eq!(out, vec!["4"]);
}

#[test]
fn builtin_val() {
    let out = run_vb(r#"
Module Program
    Sub Main()
        Console.WriteLine(Val("123") + 1)
    End Sub
End Module
"#);
    assert_eq!(out, vec!["124"]);
}

#[test]
fn builtin_isnothing() {
    let out = run_vb(r#"
Module Program
    Sub Main()
        Dim x As Object = Nothing
        Console.WriteLine(IsNothing(x))
        Dim y As Integer = 5
        Console.WriteLine(IsNothing(y))
    End Sub
End Module
"#);
    assert_eq!(out, super::helpers::dotnet_expected_lines(&["true", "false"]));
}

#[test]
fn builtin_isnumeric() {
    let out = run_vb(r#"
Module Program
    Sub Main()
        Console.WriteLine(IsNumeric("123"))
        Console.WriteLine(IsNumeric("abc"))
        Console.WriteLine(IsNumeric(42))
    End Sub
End Module
"#);
    assert_eq!(out, super::helpers::dotnet_expected_lines(&["true", "false", "true"]));
}

// ============================================================
// Array builtins
// ============================================================

#[test]
fn builtin_ubound() {
    let out = run_vb(r#"
Module Program
    Sub Main()
        Dim arr() As Integer = {10, 20, 30}
        Console.WriteLine(UBound(arr))
    End Sub
End Module
"#);
    assert_eq!(out, vec!["2"]);
}

// ============================================================
// Select Case
// ============================================================

#[test]
fn select_case_basic() {
    let out = run_vb(r#"
Module Program
    Sub Main()
        Dim x As Integer = 2
        Select Case x
            Case 1
                Console.WriteLine("one")
            Case 2
                Console.WriteLine("two")
            Case 3
                Console.WriteLine("three")
        End Select
    End Sub
End Module
"#);
    assert_eq!(out, vec!["two"]);
}

#[test]
fn select_case_else() {
    let out = run_vb(r#"
Module Program
    Sub Main()
        Dim x As Integer = 99
        Select Case x
            Case 1
                Console.WriteLine("one")
            Case 2
                Console.WriteLine("two")
            Case Else
                Console.WriteLine("other")
        End Select
    End Sub
End Module
"#);
    assert_eq!(out, vec!["other"]);
}

#[test]
fn select_case_string() {
    let out = run_vb(r#"
Module Program
    Sub Main()
        Dim color As String = "red"
        Select Case color
            Case "blue"
                Console.WriteLine("sky")
            Case "red"
                Console.WriteLine("fire")
            Case "green"
                Console.WriteLine("grass")
        End Select
    End Sub
End Module
"#);
    assert_eq!(out, vec!["fire"]);
}

// ============================================================
// Math builtins via method call syntax
// ============================================================

#[test]
fn math_methods() {
    let out = run_vb(r#"
Module Program
    Sub Main()
        Console.WriteLine(Math.Abs(-7))
        Console.WriteLine(Math.Sqrt(16))
        Console.WriteLine(Math.Pow(2, 8))
        Console.WriteLine(Math.Min(3, 7))
        Console.WriteLine(Math.Max(3, 7))
    End Sub
End Module
"#);
    assert_eq!(out, vec!["7", "4", "256", "3", "7"]);
}

// ============================================================
// ForEach
// ============================================================

#[test]
fn foreach_loop() {
    let out = run_vb(r#"
Module Program
    Sub Main()
        Dim items() As String = {"apple", "banana", "cherry"}
        For Each item As String In items
            Console.WriteLine(item)
        Next
    End Sub
End Module
"#);
    assert_eq!(out, vec!["apple", "banana", "cherry"]);
}

// ============================================================
// Nested function calls with builtins
// ============================================================

#[test]
fn nested_builtin_calls() {
    let out = run_vb(r#"
Module Program
    Sub Main()
        Console.WriteLine(UCase(Left("hello world", 5)))
        Console.WriteLine(Len(Trim("  hi  ")))
    End Sub
End Module
"#);
    assert_eq!(out, vec!["HELLO", "2"]);
}
