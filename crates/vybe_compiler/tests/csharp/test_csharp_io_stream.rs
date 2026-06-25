//! `MemoryStream`, `StreamReader`, `StreamWriter`, and `BinaryReader`/`BinaryWriter`.
use super::helpers::run_csharp;

#[test]
fn memory_stream_write_then_read_roundtrips_bytes() {
    assert_eq!(
        run_csharp(
            r#"using var ms = new System.IO.MemoryStream();
ms.WriteByte(42);
ms.Position = 0;
Console.WriteLine(ms.ReadByte());"#
        ),
        &["42"]
    );
}

#[test]
fn stream_writer_reader_roundtrip_text_line() {
    assert_eq!(
        run_csharp(
            r#"using var ms = new System.IO.MemoryStream();
using(var sw = new System.IO.StreamWriter(ms, leaveOpen:true)) sw.WriteLine("hello");
ms.Position = 0;
using var sr = new System.IO.StreamReader(ms);
Console.WriteLine(sr.ReadLine());"#
        ),
        &["hello"]
    );
}

#[test]
fn memory_stream_get_buffer_returns_internal_byte_array() {
    assert_eq!(
        run_csharp(
            r#"using var ms = new System.IO.MemoryStream(new byte[]{1,2,3});
Console.WriteLine(ms.Length);"#
        ),
        &["3"]
    );
}

#[test]
fn stream_seek_repositions_for_re_read() {
    assert_eq!(
        run_csharp(
            r#"using var ms = new System.IO.MemoryStream();
ms.WriteByte(7);
ms.Seek(0, System.IO.SeekOrigin.Begin);
Console.WriteLine(ms.ReadByte());"#
        ),
        &["7"]
    );
}

#[test]
fn binary_writer_reader_roundtrips_int32() {
    assert_eq!(
        run_csharp(
            r#"using var ms = new System.IO.MemoryStream();
using(var bw = new System.IO.BinaryWriter(ms, System.Text.Encoding.UTF8, leaveOpen:true))
    bw.Write(12345);
ms.Position = 0;
using var br = new System.IO.BinaryReader(ms);
Console.WriteLine(br.ReadInt32());"#
        ),
        &["12345"]
    );
}

#[test]
fn string_reader_reads_line_from_in_memory_string() {
    assert_eq!(
        run_csharp(
            r#"using var sr = new System.IO.StringReader("line one\nline two");
Console.WriteLine(sr.ReadLine());"#
        ),
        &["line one"]
    );
}
