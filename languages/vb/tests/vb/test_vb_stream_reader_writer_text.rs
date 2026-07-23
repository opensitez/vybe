use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: StreamReader & StreamWriter Operations
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_memory_stream_reader_writer() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Using ms As New MemoryStream()
            Using writer As New StreamWriter(ms, System.Text.Encoding.UTF8, 1024, leaveOpen:=True)
                writer.WriteLine("Line1")
                writer.WriteLine("Line2")
                writer.Flush()
            End Using

            ms.Position = 0

            Using reader As New StreamReader(ms)
                Console.WriteLine(reader.ReadLine())
                Console.WriteLine(reader.ReadLine())
            End Using
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Line1", "Line2"]);
}
