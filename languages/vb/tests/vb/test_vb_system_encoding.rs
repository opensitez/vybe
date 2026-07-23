use super::helpers::run_vb;

#[test]
fn utf8_roundtrip_roundtrips_unicode_text() {
    let out = run_vb(
        r#"
Imports System.Text

Module M
    Sub Main()
        Dim bytes As Byte() = Encoding.UTF8.GetBytes("café")
        Dim text As String = Encoding.UTF8.GetString(bytes)
        Console.WriteLine(text)
        Console.WriteLine(bytes.Length)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["café", "5"]);
}

#[test]
fn ascii_roundtrip_for_plain_ascii() {
    let out = run_vb(
        r#"
Imports System.Text

Module M
    Sub Main()
        Dim bytes As Byte() = Encoding.ASCII.GetBytes("Hello")
        Console.WriteLine(Encoding.ASCII.GetString(bytes))
        Console.WriteLine(bytes.Length)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["Hello", "5"]);
}

#[test]
fn unicode_encoding_length_changes() {
    let out = run_vb(
        r#"
Imports System.Text

Module M
    Sub Main()
        Console.WriteLine(Encoding.UTF8.GetByteCount("ab"))
        Console.WriteLine(Encoding.Unicode.GetByteCount("ab"))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["2", "4"]);
}

#[test]
fn default_encoding_is_available() {
    let out = run_vb(
        r#"
Imports System.Text

Module M
    Sub Main()
        Console.WriteLine(Encoding.Default.GetByteCount("ok"))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["2"]);
}

#[test]
fn utf16_encoding_roundtrip_rounds_trip() {
    let out = run_vb(
        r#"
Imports System.Text

Module M
    Sub Main()
        Dim bytes As Byte() = Encoding.Unicode.GetBytes("x")
        Console.WriteLine(Encoding.Unicode.GetString(bytes))
        Console.WriteLine(bytes.Length)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["x", "2"]);
}
