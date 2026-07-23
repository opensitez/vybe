use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: System.Text.Encoding (UTF8, ASCII, Unicode, UTF32)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_encoding_utf8_get_bytes_and_get_string() {
    let src = r#"
Imports System.Text

Module Program
    Sub Main()
        Dim original = "UTF8 Test"
        Dim bytes = Encoding.UTF8.GetBytes(original)
        Dim restored = Encoding.UTF8.GetString(bytes)
        Console.WriteLine(restored)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["UTF8 Test"]);
}

#[test]
fn test_vb_encoding_ascii_get_bytes_replaces_non_ascii() {
    let src = r#"
Imports System.Text

Module Program
    Sub Main()
        ' ASCII replaces non-ASCII characters with '?' (63)
        Dim text = "Hello World"
        Dim bytes = Encoding.ASCII.GetBytes(text)
        Dim restored = Encoding.ASCII.GetString(bytes)
        Console.WriteLine(restored)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Hello World"]);
}

#[test]
fn test_vb_encoding_unicode_utf16_little_endian() {
    let src = r#"
Imports System.Text

Module Program
    Sub Main()
        Dim text = "AB"
        Dim bytes = Encoding.Unicode.GetBytes(text)
        ' 'A' is 65,0 in UTF-16LE; 'B' is 66,0
        Console.WriteLine(String.Join(",", bytes))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["65,0,66,0"]);
}

#[test]
fn test_vb_encoding_big_endian_unicode_utf16_be() {
    let src = r#"
Imports System.Text

Module Program
    Sub Main()
        Dim text = "AB"
        Dim bytes = Encoding.BigEndianUnicode.GetBytes(text)
        ' 'A' is 0,65 in UTF-16BE; 'B' is 0,66
        Console.WriteLine(String.Join(",", bytes))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0,65,0,66"]);
}

#[test]
fn test_vb_encoding_utf32_get_byte_count() {
    let src = r#"
Imports System.Text

Module Program
    Sub Main()
        Dim text = "ABC" ' 3 characters
        Dim byteCount = Encoding.UTF32.GetByteCount(text)
        ' UTF-32 uses 4 bytes per character = 12 bytes
        Console.WriteLine(byteCount)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["12"]);
}

#[test]
fn test_vb_encoding_get_preamble_utf8_bom() {
    let src = r#"
Imports System.Text

Module Program
    Sub Main()
        Dim utf8Bom = Encoding.UTF8.GetPreamble()
        ' UTF-8 BOM is 239, 187, 191 (EF BB BF)
        Console.WriteLine(String.Join(",", utf8Bom))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["239,187,191"]);
}

#[test]
fn test_vb_encoding_utf8_without_bom_constructor() {
    let src = r#"
Imports System.Text

Module Program
    Sub Main()
        Dim utf8NoBom As New UTF8Encoding(False)
        Dim preamble = utf8NoBom.GetPreamble()
        Console.WriteLine(preamble.Length)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0"]);
}

#[test]
fn test_vb_encoding_latin1_iso_8859_1() {
    let src = r#"
Imports System.Text

Module Program
    Sub Main()
        Dim latin1 = Encoding.Latin1
        Dim bytes = latin1.GetBytes("Café")
        Dim restored = latin1.GetString(bytes)
        Console.WriteLine(restored)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Café"]);
}

#[test]
fn test_vb_encoding_get_char_count_buffer() {
    let src = r#"
Imports System.Text

Module Program
    Sub Main()
        Dim bytes As Byte() = Encoding.UTF8.GetBytes("VisualBasic")
        Dim charCount = Encoding.UTF8.GetCharCount(bytes, 0, 6)
        Console.WriteLine(charCount)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["6"]);
}

#[test]
fn test_vb_encoding_get_chars_array_subslice() {
    let src = r#"
Imports System.Text

Module Program
    Sub Main()
        Dim bytes As Byte() = Encoding.UTF8.GetBytes("VisualBasic")
        Dim chars(5) As Char
        Dim count = Encoding.UTF8.GetChars(bytes, 0, 6, chars, 0)
        Console.WriteLine(count & ":" & New String(chars))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["6:Visual"]);
}

#[test]
fn test_vb_encoding_get_max_byte_count() {
    let src = r#"
Imports System.Text

Module Program
    Sub Main()
        Dim maxBytes = Encoding.UTF8.GetMaxByteCount(10)
        Console.WriteLine(maxBytes >= 30)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_encoding_get_max_char_count() {
    let src = r#"
Imports System.Text

Module Program
    Sub Main()
        Dim maxChars = Encoding.UTF8.GetMaxCharCount(10)
        Console.WriteLine(maxChars >= 10)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_encoding_custom_encoder_fallback_exception() {
    let src = r#"
Imports System
Imports System.Text

Module Program
    Sub Main()
        Dim enc As Encoding = Encoding.GetEncoding("us-ascii", EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback)
        Try
            enc.GetBytes("NonAscii: €")
        Catch ex As EncoderFallbackException
            Console.WriteLine("EncoderFallbackException Caught")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["EncoderFallbackException Caught"]);
}

#[test]
fn test_vb_encoding_custom_decoder_fallback_exception() {
    let src = r#"
Imports System
Imports System.Text

Module Program
    Sub Main()
        Dim enc As Encoding = Encoding.GetEncoding("utf-8", EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback)
        Dim invalidBytes As Byte() = {&HFE, &HFF} ' Invalid UTF-8 sequence
        Try
            enc.GetString(invalidBytes)
        Catch ex As DecoderFallbackException
            Console.WriteLine("DecoderFallbackException Caught")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["DecoderFallbackException Caught"]);
}

#[test]
fn test_vb_encoding_web_name_and_header_name() {
    let src = r#"
Imports System.Text

Module Program
    Sub Main()
        Dim enc = Encoding.UTF8
        Console.WriteLine(enc.WebName & "|" & enc.HeaderName)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["utf-8|utf-8"]);
}

#[test]
fn test_vb_encoding_get_encoding_by_code_page() {
    let src = r#"
Imports System.Text

Module Program
    Sub Main()
        Dim enc = Encoding.GetEncoding(65001) ' Code page 65001 = UTF-8
        Console.WriteLine(enc.WebName)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["utf-8"]);
}

#[test]
fn test_vb_encoding_clone_customization() {
    let src = r#"
Imports System.Text

Module Program
    Sub Main()
        Dim enc As Encoding = CType(Encoding.UTF8.Clone(), Encoding)
        Console.WriteLine(enc.IsReadOnly)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False"]);
}

#[test]
fn test_vb_encoding_get_bytes_read_only_span() {
    let src = r#"
Imports System
Imports System.Text

Module Program
    Sub Main()
        Dim textSpan As ReadOnlySpan(Of Char) = "SpanEncoding".ToCharArray()
        Dim bytes = Encoding.UTF8.GetBytes(textSpan)
        Console.WriteLine(Encoding.UTF8.GetString(bytes))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["SpanEncoding"]);
}

#[test]
fn test_vb_encoding_get_string_read_only_span() {
    let src = r#"
Imports System
Imports System.Text

Module Program
    Sub Main()
        Dim byteSpan As ReadOnlySpan(Of Byte) = Encoding.UTF8.GetBytes("SpanGetString")
        Dim text = Encoding.UTF8.GetString(byteSpan)
        Console.WriteLine(text)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["SpanGetString"]);
}

#[test]
fn test_vb_encoding_equality_comparison() {
    let src = r#"
Imports System.Text

Module Program
    Sub Main()
        Dim e1 = Encoding.UTF8
        Dim e2 = Encoding.GetEncoding("utf-8")
        Console.WriteLine(e1.Equals(e2))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}
