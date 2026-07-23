use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: System.BitConverter Operations
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_bit_converter_get_bytes_to_int32() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim val As Integer = 123456789
        Dim bytes As Byte() = BitConverter.GetBytes(val)
        Console.WriteLine(bytes.Length)
        Dim restored As Integer = BitConverter.ToInt32(bytes, 0)
        Console.WriteLine(restored)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["4", "123456789"]);
}
