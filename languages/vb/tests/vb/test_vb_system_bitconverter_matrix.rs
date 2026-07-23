use super::helpers::run_vb;

#[test]
fn bitconverter_roundtrip_int32_bytes() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim source As Integer = 123456
        Dim bytes() As Byte = BitConverter.GetBytes(source)
        Console.WriteLine(BitConverter.ToInt32(bytes, 0))
        Console.WriteLine(bytes.Length)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["123456", "4"]);
}

#[test]
fn bitconverter_roundtrip_double_bytes() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim source As Double = 12.5
        Dim bytes() As Byte = BitConverter.GetBytes(source)
        Console.WriteLine(BitConverter.ToDouble(bytes, 0))
        Console.WriteLine(bytes.Length)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["12.5", "8"]);
}

#[test]
fn bitconverter_works_with_boolean_bytes() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim bytes() As Byte = BitConverter.GetBytes(True)
        Console.WriteLine(bytes(0))
        Console.WriteLine(bytes.Length)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["1", "1"]);
}

#[test]
fn bitconverter_is_little_endian_flag_is_stable_boolean() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Console.WriteLine(BitConverter.IsLittleEndian)
    End Sub
End Module
"#,
    );

    assert_eq!(out.len(), 1);
    assert!(out[0] == "True" || out[0] == "False");
}

#[test]
fn bitconverter_tohex_preserves_byte_count() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim bytes() As Byte = {&H01, &H2A, &HFF}
        Dim text As String = BitConverter.ToString(bytes)
        Console.WriteLine(text.Length)
        Console.WriteLine(text.StartsWith("01-2A"))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["8", "True"]);
}

#[test]
fn bitconverter_to_int64_roundtrip() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim source As Long = 9876543210L
        Dim bytes() As Byte = BitConverter.GetBytes(source)
        Dim restored As Long = BitConverter.ToInt64(bytes, 0)
        Console.WriteLine(bytes.Length)
        Console.WriteLine(restored)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["8", "9876543210"]);
}

#[test]
fn bitconverter_to_string_has_dash_delimiters() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim bytes() As Byte = {&H0A, &H0B}
        Console.WriteLine(BitConverter.ToString(bytes).Contains("-"))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}

#[test]
fn bitconverter_to_unicode_string_from_bytes() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim bytes() As Byte = {&H41, &H00}
        Dim value As Char = BitConverter.ToChar(bytes, 0)
        Console.WriteLine(value)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["A"]);
}

#[test]
fn bitconverter_copy_from_readonly_span_style_bytes() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim source As Integer = 255
        Dim bytes() As Byte = New Byte(3) {}
        Buffer.BlockCopy(BitConverter.GetBytes(source), 0, bytes, 0, 4)
        Console.WriteLine(bytes(0))
        Console.WriteLine(bytes(3))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["255", "0"]);
}
