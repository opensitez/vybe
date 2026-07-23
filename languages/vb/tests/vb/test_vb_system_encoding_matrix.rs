use super::helpers::run_vb;

#[test]
fn encoding_utf8_byte_count_for_ascii() {
    let out = run_vb(
        r#"
Imports System
Imports System.Text

Module M
    Sub Main()
        Dim bytes() As Byte = Encoding.UTF8.GetBytes("hello")
        Console.WriteLine(bytes.Length)
        Console.WriteLine(Encoding.UTF8.GetByteCount("hello"))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["5", "5"]);
}

#[test]
fn encoding_utf8_roundtrips_unicode() {
    let out = run_vb(
        r#"
Imports System
Imports System.Text

Module M
    Sub Main()
        Dim text As String = "café"
        Dim bytes() As Byte = Encoding.UTF8.GetBytes(text)
        Dim decoded As String = Encoding.UTF8.GetString(bytes)

        Console.WriteLine(text = decoded)
        Console.WriteLine(bytes.Length > 4)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn encoding_ascii_roundtrip_ascii_text_only() {
    let out = run_vb(
        r#"
Imports System
Imports System.Text

Module M
    Sub Main()
        Dim bytes() As Byte = Encoding.ASCII.GetBytes("ABC")
        Dim text As String = Encoding.ASCII.GetString(bytes)

        Console.WriteLine(text)
        Console.WriteLine(bytes(0))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["ABC", "65"]);
}

#[test]
fn encoding_unicode_uses_two_bytes_for_ascii_char() {
    let out = run_vb(
        r#"
Imports System
Imports System.Text

Module M
    Sub Main()
        Dim bytes() As Byte = Encoding.Unicode.GetBytes("x")
        Console.WriteLine(bytes.Length)
        Console.WriteLine(Encoding.Unicode.GetString(bytes))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["2", "x"]);
}

#[test]
fn encoding_default_roundtrip() {
    let out = run_vb(
        r#"
Imports System
Imports System.Text

Module M
    Sub Main()
        Dim text As String = "runtime"
        Dim bytes() As Byte = Encoding.Default.GetBytes(text)
        Dim restored As String = Encoding.Default.GetString(bytes)

        Console.WriteLine(restored)
        Console.WriteLine(bytes.Length)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["runtime", "7"]);
}

#[test]
fn encoding_utf32_count_matches_formula() {
    let out = run_vb(
        r#"
Imports System
Imports System.Text

Module M
    Sub Main()
        Dim bytes() As Byte = Encoding.UTF32.GetBytes("A")
        Console.WriteLine(bytes.Length)
        Console.WriteLine(Encoding.UTF32.GetString(bytes))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["4", "A"]);
}

#[test]
fn encoding_preamble_length_is_stable_for_utf8() {
    let out = run_vb(
        r#"
Imports System
Imports System.Text

Module M
    Sub Main()
        Dim preamble() As Byte = Encoding.UTF8.GetPreamble()
        Console.WriteLine(preamble.Length >= 0)
        Console.WriteLine(preamble.Length = 3 OrElse preamble.Length = 0)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn encoding_convert_between_unicode_and_utf8_roundtrip() {
    let out = run_vb(
        r#"
Imports System
Imports System.Text

Module M
    Sub Main()
        Dim src() As Byte = Encoding.Unicode.GetBytes("vb")
        Dim utf8() As Byte = Encoding.Convert(Encoding.Unicode, Encoding.UTF8, src)
        Dim back() As Byte = Encoding.Convert(Encoding.UTF8, Encoding.Unicode, utf8)

        Console.WriteLine(Encoding.Unicode.GetString(back))
        Console.WriteLine(utf8.Length > 0)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["vb", "True"]);
}

#[test]
fn encoding_max_byte_count_is_non_negative_for_ascii() {
    let out = run_vb(
        r#"
Imports System
Imports System.Text

Module M
    Sub Main()
        Console.WriteLine(Encoding.UTF8.GetMaxByteCount(0) >= 0)
        Console.WriteLine(Encoding.Unicode.GetMaxByteCount(1))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "2"]);
}

#[test]
fn encoding_getbytes_and_getstring_are_inverse_for_emoji() {
    let out = run_vb(
        r#"
Imports System
Imports System.Text

Module M
    Sub Main()
        Dim emoji As String = "😀"
        Dim bytes() As Byte = Encoding.UTF8.GetBytes(emoji)
        Dim text As String = Encoding.UTF8.GetString(bytes)

        Console.WriteLine(text = emoji)
        Console.WriteLine(bytes.Length >= 4)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}
