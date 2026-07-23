use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: System.Convert.ToBase64String & FromBase64String
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_convert_to_base64_string_basic() {
    let src = r#"
Imports System
Imports System.Text

Module Program
    Sub Main()
        Dim bytes As Byte() = Encoding.UTF8.GetBytes("Hello World")
        Dim base64Str = Convert.ToBase64String(bytes)
        Console.WriteLine(base64Str)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["SGVsbG8gV29ybGQ="]);
}

#[test]
fn test_vb_convert_from_base64_string_roundtrip() {
    let src = r#"
Imports System
Imports System.Text

Module Program
    Sub Main()
        Dim b64 = "SGVsbG8gV29ybGQ="
        Dim bytes As Byte() = Convert.FromBase64String(b64)
        Dim str = Encoding.UTF8.GetString(bytes)
        Console.WriteLine(str)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Hello World"]);
}

#[test]
fn test_vb_convert_to_base64_string_offset_and_length() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim raw As Byte() = {10, 20, 30, 40, 50}
        ' Convert slice starting at index 1 for length 3
        Dim b64 = Convert.ToBase64String(raw, 1, 3)
        Dim restored As Byte() = Convert.FromBase64String(b64)
        Console.WriteLine(String.Join(",", restored))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["20,30,40"]);
}

#[test]
fn test_vb_convert_to_base64_string_options_insert_line_breaks() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim largePayload As Byte() = New Byte(99) {}
        For i As Integer = 0 To 99 : largePayload(i) = CByte(i) : Next
        Dim b64Formatted = Convert.ToBase64String(largePayload, Base64FormattingOptions.InsertLineBreaks)
        Console.WriteLine(b64Formatted.Contains(vbCrLf) OrElse b64Formatted.Contains(vbLf))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_convert_try_from_base64_string_chars() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim chars As Char() = "SGVsbG8=".ToCharArray()
        Dim buffer(9) As Byte
        Dim bytesWritten As Integer
        Dim ok = Convert.TryFromBase64Chars(chars, buffer, bytesWritten)
        Console.WriteLine(ok & "|" & bytesWritten)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|5"]);
}

#[test]
fn test_vb_convert_to_base64_char_array() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim input As Byte() = {1, 2, 3}
        Dim outChars(10) As Char
        ' ToBase64CharArray(inArray, offset, length, outArray, outOffset)
        Dim count = Convert.ToBase64CharArray(input, 0, 3, outChars, 0)
        Console.WriteLine(count & ":" & New String(outChars, 0, count))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["4:AQID"]);
}

#[test]
fn test_vb_convert_from_base64_invalid_format_throws() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Try
            Convert.FromBase64String("Invalid!Base64@String")
        Catch ex As FormatException
            Console.WriteLine("FormatException Caught on Invalid Base64")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["FormatException Caught on Invalid Base64"]
    );
}

#[test]
fn test_vb_convert_from_base64_null_throws_argument_null() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Try
            Convert.FromBase64String(Nothing)
        Catch ex As ArgumentNullException
            Console.WriteLine("ArgumentNullException Caught on Null Base64")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["ArgumentNullException Caught on Null Base64"]
    );
}

#[test]
fn test_vb_convert_to_base64_empty_array_returns_empty_string() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim b64 = Convert.ToBase64String(New Byte() {})
        Console.WriteLine(b64 = "")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_convert_from_base64_whitespace_ignored() {
    let src = r#"
Imports System
Imports System.Text

Module Program
    Sub Main()
        Dim b64WithSpaces = " SGVs bG8g V29ybGQ= "
        Dim bytes = Convert.FromBase64String(b64WithSpaces)
        Console.WriteLine(Encoding.UTF8.GetString(bytes))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Hello World"]);
}

#[test]
fn test_vb_convert_to_base64_url_safe_replacement_simulation() {
    let src = r#"
Imports System
Imports System.Text

Module Program
    Private Function ToBase64Url(data As Byte()) As String
        Dim base64 = Convert.ToBase64String(data)
        Return base64.Replace("+", "-").Replace("/", "_").TrimEnd("="c)
    End Function

    Sub Main()
        Dim bytes As Byte() = Encoding.UTF8.GetBytes("Subject?Data#1")
        Dim urlSafe = ToBase64Url(bytes)
        Console.WriteLine(Not urlSafe.Contains("+") AndAlso Not urlSafe.Contains("/"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_convert_from_base64_url_safe_restoration() {
    let src = r#"
Imports System
Imports System.Text

Module Program
    Private Function FromBase64Url(base64Url As String) As Byte()
        Dim padded = base64Url.Replace("-", "+").Replace("_", "/")
        Select Case padded.Length Mod 4
            Case 2 : padded &= "=="
            Case 3 : padded &= "="
        End Select
        Return Convert.FromBase64String(padded)
    End Function

    Sub Main()
        Dim bytes = FromBase64Url("U3ViamVjdD9EYXRhIzE")
        Console.WriteLine(Encoding.UTF8.GetString(bytes))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Subject?Data#1"]);
}

#[test]
fn test_vb_convert_to_base64_binary_struct_serialization() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

<StructLayout(LayoutKind.Sequential)>
Structure RecordHeader
    Public Magic As Integer
    Public Length As Short
End Structure

Module Program
    Sub Main()
        Dim h As New RecordHeader With {.Magic = &H41424344, .Length = 100}
        Dim size = Marshal.SizeOf(GetType(RecordHeader))
        Dim ptr = Marshal.AllocHGlobal(size)
        Marshal.StructureToPtr(h, ptr, False)

        Dim bytes(size - 1) As Byte
        Marshal.Copy(ptr, bytes, 0, size)
        Marshal.FreeHGlobal(ptr)

        Dim b64 = Convert.ToBase64String(bytes)
        Console.WriteLine(b64.Length > 0)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_convert_to_base64_chunked_stream_reading() {
    let src = r#"
Imports System
Imports System.IO

Module Program
    Sub Main()
        Dim data As Byte() = Encoding.UTF8.GetBytes("Chunk1Chunk2Chunk3")
        Using ms As New MemoryStream(data)
            Dim buffer(5) As Byte
            Dim readCount = ms.Read(buffer, 0, 6)
            Dim chunkB64 = Convert.ToBase64String(buffer, 0, readCount)
            Console.WriteLine(chunkB64)
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Q2h1bmsx"]);
}

#[test]
fn test_vb_convert_from_base64_span_buffer() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim span As ReadOnlySpan(Of Char) = "AQID".ToCharArray()
        Dim dest(3) As Byte
        Dim bytesWritten As Integer
        Dim ok = Convert.TryFromBase64Chars(span, dest, bytesWritten)
        Console.WriteLine(ok & ":" & String.Join(",", dest, 0, bytesWritten))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True:1,2,3"]);
}

#[test]
fn test_vb_convert_to_base64_span_buffer() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim srcBytes As ReadOnlySpan(Of Byte) = New Byte() {1, 2, 3}
        Dim b64Str = Convert.ToBase64String(srcBytes)
        Console.WriteLine(b64Str)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["AQID"]);
}

#[test]
fn test_vb_convert_to_base64_string_null_array_throws() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Try
            Convert.ToBase64String(Nothing)
        Catch ex As ArgumentNullException
            Console.WriteLine("ArgumentNullException Caught on Null Byte Array")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["ArgumentNullException Caught on Null Byte Array"]
    );
}

#[test]
fn test_vb_convert_to_base64_out_of_range_length_throws() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim data As Byte() = {1, 2, 3}
        Try
            Convert.ToBase64String(data, 0, 10)
        Catch ex As ArgumentOutOfRangeException
            Console.WriteLine("ArgumentOutOfRangeException Caught on Invalid Length")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["ArgumentOutOfRangeException Caught on Invalid Length"]
    );
}

#[test]
fn test_vb_convert_to_base64_single_byte_padding_check() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        ' Single byte input produces 4-char base64 string with '==' padding!
        Dim b64 = Convert.ToBase64String(New Byte() {65})
        Console.WriteLine(b64 & "|EndsWith==" & b64.EndsWith("=="))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["QQ==|EndsWith==True"]);
}

#[test]
fn test_vb_convert_to_base64_two_bytes_padding_check() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        ' Two byte input produces 4-char base64 string with '=' padding!
        Dim b64 = Convert.ToBase64String(New Byte() {65, 66})
        Console.WriteLine(b64 & "|EndsWith=" & b64.EndsWith("="))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["QUI=|EndsWith=True"]);
}
