use super::helpers::run_vb;

#[test]
fn memory_stream_read_write_text() {
    let out = run_vb(
        r#"
Imports System
Imports System.IO

Module M
    Sub Main()
        Using ms As New MemoryStream()
            Dim writer As New StreamWriter(ms)
            writer.Write("abc")
            writer.Flush()
            Console.WriteLine(ms.Length > 0)
            ms.Position = 0
            Using reader As New StreamReader(ms)
                Console.WriteLine(reader.ReadToEnd())
            End Using
        End Using
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "abc"]);
}

#[test]
fn memory_stream_to_array_roundtrip() {
    let out = run_vb(
        r#"
Imports System.IO

Module M
    Sub Main()
        Dim bytes() As Byte = {1, 2, 3}
        Using ms As New MemoryStream(bytes)
            Dim cloned As Byte() = ms.ToArray()
            Console.WriteLine(cloned.Length)
            Console.WriteLine(cloned(0))
            Console.WriteLine(cloned(2))
        End Using
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["3", "1", "3"]);
}

#[test]
fn binary_writer_reads_back_different_types() {
    let out = run_vb(
        r#"
Imports System.IO

Module M
    Sub Main()
        Using ms As New MemoryStream()
            Using writer As New BinaryWriter(ms)
                writer.Write(123)
                writer.Write(True)
                writer.Write("done")
            End Using
            ms.Position = 0
            Using reader As New BinaryReader(ms)
                Console.WriteLine(reader.ReadInt32())
                Console.WriteLine(reader.ReadBoolean())
                Console.WriteLine(reader.ReadString())
            End Using
        End Using
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["123", "True", "done"]);
}

#[test]
fn stream_copy_to_memory_roundtrip() {
    let out = run_vb(
        r#"
Imports System.IO
Imports System.Text

Module M
    Sub Main()
        Dim source As New MemoryStream(Encoding.UTF8.GetBytes("copy-source"))
        Dim destination As New MemoryStream()
        source.CopyTo(destination)
        Console.WriteLine(destination.Length)
        Console.WriteLine(Encoding.UTF8.GetString(destination.ToArray()))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["11", "copy-source"]);
}

#[test]
fn stream_seek_and_position_workflow() {
    let out = run_vb(
        r#"
Imports System.IO

Module M
    Sub Main()
        Dim bytes() As Byte = {10, 20, 30, 40}
        Using ms As New MemoryStream(bytes)
            ms.Seek(2, SeekOrigin.Begin)
            Console.WriteLine(ms.Position)
            Dim one As Integer = ms.ReadByte()
            Console.WriteLine(one)
            Console.WriteLine(ms.Position)
        End Using
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["2", "30", "3"]);
}

#[test]
fn stream_reader_reads_line_by_line() {
    let out = run_vb(
        r#"
Imports System.IO
Imports System.Text

Module M
    Sub Main()
        Using ms As New MemoryStream(Encoding.UTF8.GetBytes("one" & vbLf & "two"))
            Using reader As New StreamReader(ms)
                Console.WriteLine(reader.ReadLine())
                Console.WriteLine(reader.ReadLine())
            End Using
        End Using
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["one", "two"]);
}
