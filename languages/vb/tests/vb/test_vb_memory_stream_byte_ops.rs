use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: MemoryStream Byte Read/Write Operations
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_memory_stream_write_to_array() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Using ms As New MemoryStream()
            ms.WriteByte(65) ' 'A'
            ms.WriteByte(66) ' 'B'
            Dim bytes As Byte() = ms.ToArray()
            Console.WriteLine(bytes.Length)
            Console.WriteLine(bytes(0) & ":" & bytes(1))
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2", "65:66"]);
}
