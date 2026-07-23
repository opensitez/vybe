use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Advanced Numeric Type Conversions
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_conv_byte_to_integer() {
    let src = r#"
Module Program
    Sub Main()
        Dim b As Byte = 255
        Dim i As Integer = CInt(b)
        Console.WriteLine(i)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["255"]);
}

#[test]
fn test_vb_conv_sbyte_to_short() {
    let src = r#"
Module Program
    Sub Main()
        Dim sb As SByte = -128
        Dim s As Short = CShort(sb)
        Console.WriteLine(s)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["-128"]);
}

#[test]
fn test_vb_conv_uinteger_to_ulong() {
    let src = r#"
Module Program
    Sub Main()
        Dim ui As UInteger = 4294967295UI
        Dim ul As ULong = CULng(ui)
        Console.WriteLine(ul)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["4294967295"]);
}

#[test]
fn test_vb_conv_double_to_integer_banker_rounding() {
    let src = r#"
Module Program
    Sub Main()
        Dim d1 As Double = 2.5
        Dim d2 As Double = 3.5
        Console.WriteLine(CInt(d1))
        Console.WriteLine(CInt(d2))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2", "4"]);
}

#[test]
fn test_vb_conv_single_to_decimal() {
    let src = r#"
Module Program
    Sub Main()
        Dim f As Single = 12.5F
        Dim dec As Decimal = CDec(f)
        Console.WriteLine(dec)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["12.5"]);
}

#[test]
fn test_vb_conv_string_hex_to_integer() {
    let src = r#"
Module Program
    Sub Main()
        Dim val As Integer = Convert.ToInt32("FF", 16)
        Console.WriteLine(val)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["255"]);
}

#[test]
fn test_vb_conv_string_octal_to_integer() {
    let src = r#"
Module Program
    Sub Main()
        Dim val As Integer = Convert.ToInt32("77", 8)
        Console.WriteLine(val)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["63"]);
}

#[test]
fn test_vb_conv_string_binary_to_integer() {
    let src = r#"
Module Program
    Sub Main()
        Dim val As Integer = Convert.ToInt32("101010", 2)
        Console.WriteLine(val)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["42"]);
}

#[test]
fn test_vb_conv_integer_to_hex_string() {
    let src = r#"
Module Program
    Sub Main()
        Dim i As Integer = 255
        Console.WriteLine(Hex(i))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["FF"]);
}

#[test]
fn test_vb_conv_integer_to_oct_string() {
    let src = r#"
Module Program
    Sub Main()
        Dim i As Integer = 63
        Console.WriteLine(Oct(i))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["77"]);
}

#[test]
fn test_vb_conv_tryparse_integer_valid() {
    let src = r#"
Module Program
    Sub Main()
        Dim result As Integer
        Dim success As Boolean = Integer.TryParse("12345", result)
        Console.WriteLine(success)
        Console.WriteLine(result)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "12345"]);
}

#[test]
fn test_vb_conv_tryparse_integer_invalid() {
    let src = r#"
Module Program
    Sub Main()
        Dim result As Integer
        Dim success As Boolean = Integer.TryParse("abc", result)
        Console.WriteLine(success)
        Console.WriteLine(result)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False", "0"]);
}

#[test]
fn test_vb_conv_tryparse_double_scientific() {
    let src = r#"
Module Program
    Sub Main()
        Dim result As Double
        Dim success As Boolean = Double.TryParse("1.23e4", result)
        Console.WriteLine(success)
        Console.WriteLine(result)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "12300"]);
}

#[test]
fn test_vb_conv_type_code_enumeration() {
    let src = r#"
Module Program
    Sub Main()
        Dim i As Integer = 100
        Dim tc As TypeCode = Convert.GetTypeCode(i)
        Console.WriteLine(tc.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Int32"]);
}

#[test]
fn test_vb_conv_convert_change_type() {
    let src = r#"
Module Program
    Sub Main()
        Dim obj As Object = "99"
        Dim res As Object = Convert.ChangeType(obj, GetType(Integer))
        Console.WriteLine(res.GetType().Name)
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Int32", "99"]);
}

#[test]
fn test_vb_conv_char_to_uint16() {
    let src = r#"
Module Program
    Sub Main()
        Dim c As Char = "A"c
        Dim code As UShort = Convert.ToUInt16(c)
        Console.WriteLine(code)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["65"]);
}

#[test]
fn test_vb_conv_uint16_to_char() {
    let src = r#"
Module Program
    Sub Main()
        Dim code As UShort = 66
        Dim c As Char = Convert.ToChar(code)
        Console.WriteLine(c)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["B"]);
}

#[test]
fn test_vb_conv_boolean_to_integer() {
    let src = r#"
Module Program
    Sub Main()
        Dim t As Boolean = True
        Dim f As Boolean = False
        Console.WriteLine(CInt(t))
        Console.WriteLine(CInt(f))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["-1", "0"]);
}

#[test]
fn test_vb_conv_integer_to_boolean() {
    let src = r#"
Module Program
    Sub Main()
        Dim zero As Integer = 0
        Dim nonZero As Integer = -5
        Console.WriteLine(CBool(zero))
        Console.WriteLine(CBool(nonZero))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False", "True"]);
}

#[test]
fn test_vb_conv_enum_to_underlying_and_back() {
    let src = r#"
Enum ColorCode As Byte
    Red = 1
    Green = 2
    Blue = 3
End Enum

Module Program
    Sub Main()
        Dim c As ColorCode = ColorCode.Green
        Dim b As Byte = CByte(c)
        Dim c2 As ColorCode = CType(3, ColorCode)
        Console.WriteLine(b)
        Console.WriteLine(c2.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2", "Blue"]);
}
