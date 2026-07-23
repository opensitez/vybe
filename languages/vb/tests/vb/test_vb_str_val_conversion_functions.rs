use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Legacy Conversion Functions (Str, Val, Fix, Int, Oct, Hex)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_str_function_positive_number_leading_space() {
    let src = r#"
Imports Microsoft.VisualBasic

Module Program
    Sub Main()
        ' Str(positive) includes a leading space for sign!
        Dim s = Str(42)
        Console.WriteLine("[" & s & "]")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["[ 42]"]);
}

#[test]
fn test_vb_str_function_negative_number() {
    let src = r#"
Imports Microsoft.VisualBasic

Module Program
    Sub Main()
        ' Str(negative) includes minus sign without extra leading space
        Dim s = Str(-42)
        Console.WriteLine("[" & s & "]")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["[-42]"]);
}

#[test]
fn test_vb_val_function_parses_leading_digits() {
    let src = r#"
Imports Microsoft.VisualBasic

Module Program
    Sub Main()
        ' Val parses numbers until it encounters non-numeric char
        Dim v = Val("12345ABC")
        Console.WriteLine(v)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["12345"]);
}

#[test]
fn test_vb_val_function_strips_spaces() {
    let src = r#"
Imports Microsoft.VisualBasic

Module Program
    Sub Main()
        ' Val ignores whitespace inside string! "1 2 3" -> 123
        Dim v = Val(" 1 2 3 ")
        Console.WriteLine(v)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["123"]);
}

#[test]
fn test_vb_val_function_floating_point_dot() {
    let src = r#"
Imports Microsoft.VisualBasic

Module Program
    Sub Main()
        Dim v = Val("12.345.67")
        Console.WriteLine(v)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["12.345"]);
}

#[test]
fn test_vb_val_function_hex_octal_prefix() {
    let src = r#"
Imports Microsoft.VisualBasic

Module Program
    Sub Main()
        ' Val recognizes &H for Hex and &O for Octal!
        Dim vHex = Val("&HFF")
        Dim vOct = Val("&O77")
        Console.WriteLine(vHex & "|" & vOct)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["255|63"]);
}

#[test]
fn test_vb_val_function_non_numeric_returns_zero() {
    let src = r#"
Imports Microsoft.VisualBasic

Module Program
    Sub Main()
        Dim v = Val("NoDigitsHere")
        Console.WriteLine(v)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0"]);
}

#[test]
fn test_vb_hex_function_integer_formatting() {
    let src = r#"
Imports Microsoft.VisualBasic

Module Program
    Sub Main()
        Console.WriteLine(Hex(255) & "|" & Hex(16))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["FF|10"]);
}

#[test]
fn test_vb_oct_function_integer_formatting() {
    let src = r#"
Imports Microsoft.VisualBasic

Module Program
    Sub Main()
        Console.WriteLine(Oct(63) & "|" & Oct(8))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["77|10"]);
}

#[test]
fn test_vb_fix_vs_int_positive_number() {
    let src = r#"
Imports Microsoft.VisualBasic

Module Program
    Sub Main()
        Console.WriteLine(Fix(99.9) & "|" & Int(99.9))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["99|99"]);
}

#[test]
fn test_vb_fix_vs_int_negative_number() {
    let src = r#"
Imports Microsoft.VisualBasic

Module Program
    Sub Main()
        ' Fix truncates toward zero (-99); Int rounds down to lower integer (-100)!
        Console.WriteLine(Fix(-99.1) & "|" & Int(-99.1))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["-99|-100"]);
}

#[test]
fn test_vb_cbool_cbyte_cchar_cdate_cdbl_cdec_cint_clng_cobj_csbyte_cshort_csng_cstr_cuint_culng_cushort()
 {
    let src = r#"
Module Program
    Sub Main()
        Dim n = CInt("123")
        Dim d = CDbl("45.67")
        Dim s = CStr(999)
        Dim b = CBool(1)
        Console.WriteLine(n & "|" & d & "|" & s & "|" & b)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["123|45.67|999|True"]);
}

#[test]
fn test_vb_asc_and_chr_legacy_functions() {
    let src = r#"
Imports Microsoft.VisualBasic

Module Program
    Sub Main()
        Dim code = Asc("A")
        Dim ch = Chr(65)
        Console.WriteLine(code & "|" & ch)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["65|A"]);
}

#[test]
fn test_vb_val_function_exponential_notation() {
    let src = r#"
Imports Microsoft.VisualBasic

Module Program
    Sub Main()
        Dim v = Val("1.23E4")
        Console.WriteLine(v)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["12300"]);
}

#[test]
fn test_vb_str_function_double_formatting() {
    let src = r#"
Imports Microsoft.VisualBasic

Module Program
    Sub Main()
        Dim s = Str(3.14159)
        Console.WriteLine(s.Trim())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3.14159"]);
}

#[test]
fn test_vb_val_null_string_returns_zero() {
    let src = r#"
Imports Microsoft.VisualBasic

Module Program
    Sub Main()
        Dim v = Val(Nothing)
        Console.WriteLine(v)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0"]);
}

#[test]
fn test_vb_hex_function_long_formatting() {
    let src = r#"
Imports Microsoft.VisualBasic

Module Program
    Sub Main()
        Dim h = Hex(4294967295L)
        Console.WriteLine(h)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["FFFFFFFF"]);
}

#[test]
fn test_vb_oct_function_long_formatting() {
    let src = r#"
Imports Microsoft.VisualBasic

Module Program
    Sub Main()
        Dim o = Oct(511L)
        Console.WriteLine(o)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["777"]);
}

#[test]
fn test_vb_ctype_custom_conversion_expression() {
    let src = r#"
Module Program
    Sub Main()
        Dim obj As Object = "55"
        Dim num As Integer = CType(obj, Integer)
        Console.WriteLine(num)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["55"]);
}

#[test]
fn test_vb_val_and_str_roundtrip_simulation() {
    let src = r#"
Imports Microsoft.VisualBasic

Module Program
    Sub Main()
        Dim orig As Double = 987.65
        Dim s = Str(orig)
        Dim restored = Val(s)
        Console.WriteLine(orig = restored)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}
