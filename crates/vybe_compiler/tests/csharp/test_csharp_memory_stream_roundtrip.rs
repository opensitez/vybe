//! `MemoryStream` holds bytes in memory for write-then-read without files.
use super::helpers::run_csharp;

#[test]
fn memory_stream_write_then_read_returns_same_bytes() {
    assert_eq!(
        run_csharp(
            r#"
using System.IO;
var stream = new MemoryStream();
var writer = new StreamWriter(stream);
writer.Write("payload");
writer.Flush();
stream.Position = 0;
var reader = new StreamReader(stream);
Console.WriteLine(reader.ReadToEnd());
"#
        ),
        &["payload"]
    );
}

#[test]
fn memory_stream_to_array_captures_written_length_not_capacity() {
    assert_eq!(
        run_csharp(
            r#"
using System.IO;
var stream = new MemoryStream();
stream.WriteByte(1);
stream.WriteByte(2);
var bytes = stream.ToArray();
Console.WriteLine(bytes.Length);
Console.WriteLine(bytes[1]);
"#
        ),
        &["2", "2"]
    );
}

#[test]
fn memory_stream_seek_begin_repositions_read_cursor_for_second_pass() {
    assert_eq!(
        run_csharp(
            r#"
using System.IO;
var stream = new MemoryStream();
var writer = new StreamWriter(stream);
writer.Write("ab");
writer.Flush();
stream.Seek(0, SeekOrigin.Begin);
Console.WriteLine(stream.ReadByte());
stream.Seek(0, SeekOrigin.Begin);
Console.WriteLine(stream.ReadByte());
"#
        ),
        &["97", "97"]
    );
}
