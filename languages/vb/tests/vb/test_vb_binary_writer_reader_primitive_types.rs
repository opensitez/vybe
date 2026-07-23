use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: System.IO.BinaryWriter & BinaryReader Serialization
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_binary_writer_reader_roundtrip_primitives() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Using ms As New MemoryStream()
            Using bw As New BinaryWriter(ms, System.Text.Encoding.UTF8, True)
                bw.Write(42)
                bw.Write("BinaryString")
                bw.Write(3.14159)
                bw.Write(True)
            End Using

            ms.Position = 0
            Using br As New BinaryReader(ms)
                Dim num = br.ReadInt32()
                Dim str = br.ReadString()
                Dim dbl = br.ReadDouble()
                Dim flag = br.ReadBoolean()
                Console.WriteLine(num & "|" & str & "|" & dbl & "|" & flag)
            End Using
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["42|BinaryString|3.14159|True"]);
}

#[test]
fn test_vb_binary_writer_reader_byte_array() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Dim payload As Byte() = {10, 20, 30, 40, 50}
        Using ms As New MemoryStream()
            Using bw As New BinaryWriter(ms, System.Text.Encoding.UTF8, True)
                bw.Write(payload.Length)
                bw.Write(payload)
            End Using

            ms.Position = 0
            Using br As New BinaryReader(ms)
                Dim len = br.ReadInt32()
                Dim readBytes = br.ReadBytes(len)
                Console.WriteLine(String.Join(",", readBytes))
            End Using
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10,20,30,40,50"]);
}

#[test]
fn test_vb_binary_writer_reader_char_and_string() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Using ms As New MemoryStream()
            Using bw As New BinaryWriter(ms, System.Text.Encoding.UTF8, True)
                bw.Write("A"c)
                bw.Write("Hello")
            End Using

            ms.Position = 0
            Using br As New BinaryReader(ms)
                Dim ch = br.ReadChar()
                Dim str = br.ReadString()
                Console.WriteLine(ch & "|" & str)
            End Using
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["A|Hello"]);
}

#[test]
fn test_vb_binary_writer_reader_7bit_encoded_int() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Using ms As New MemoryStream()
            Using bw As New BinaryWriter(ms, System.Text.Encoding.UTF8, True)
                bw.Write7BitEncodedInt(16384)
            End Using

            ms.Position = 0
            Using br As New BinaryReader(ms)
                Dim val = br.Read7BitEncodedInt()
                Console.WriteLine(val)
            End Using
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["16384"]);
}

#[test]
fn test_vb_binary_writer_reader_decimal() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Using ms As New MemoryStream()
            Using bw As New BinaryWriter(ms, System.Text.Encoding.UTF8, True)
                bw.Write(9999999.99D)
            End Using

            ms.Position = 0
            Using br As New BinaryReader(ms)
                Dim dec = br.ReadDecimal()
                Console.WriteLine(dec)
            End Using
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["9999999.99"]);
}

#[test]
fn test_vb_binary_writer_seek_overwrite() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Using ms As New MemoryStream()
            Using bw As New BinaryWriter(ms, System.Text.Encoding.UTF8, True)
                bw.Write(100)
                bw.Write(200)
                bw.Seek(0, SeekOrigin.Begin)
                bw.Write(999)
            End Using

            ms.Position = 0
            Using br As New BinaryReader(ms)
                Dim first = br.ReadInt32()
                Dim second = br.ReadInt32()
                Console.WriteLine(first & "|" & second)
            End Using
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["999|200"]);
}

#[test]
fn test_vb_binary_reader_peek_char() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Using ms As New MemoryStream()
            Using bw As New BinaryWriter(ms, System.Text.Encoding.UTF8, True)
                bw.Write("X"c)
            End Using

            ms.Position = 0
            Using br As New BinaryReader(ms)
                Dim peeked = br.PeekChar()
                Dim readCh = br.ReadChar()
                Console.WriteLine(ChrW(peeked) & "=" & readCh)
            End Using
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["X=X"]);
}

#[test]
fn test_vb_binary_reader_end_of_stream_exception() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Using ms As New MemoryStream()
            Using br As New BinaryReader(ms)
                Try
                    Dim n = br.ReadInt32()
                Catch ex As EndOfStreamException
                    Console.WriteLine("EndOfStreamException Caught")
                End Try
            End Using
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["EndOfStreamException Caught"]);
}

#[test]
fn test_vb_binary_writer_reader_uint16_uint32_uint64() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Using ms As New MemoryStream()
            Using bw As New BinaryWriter(ms, System.Text.Encoding.UTF8, True)
                bw.Write(65000US)
                bw.Write(4000000000UI)
                bw.Write(18000000000000000000UL)
            End Using

            ms.Position = 0
            Using br As New BinaryReader(ms)
                Dim u16 = br.ReadUInt16()
                Dim u32 = br.ReadUInt32()
                Dim u64 = br.ReadUInt64()
                Console.WriteLine(u16 & "|" & u32 & "|" & u64)
            End Using
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["65000|4000000000|18000000000000000000"]);
}

#[test]
fn test_vb_binary_writer_reader_half_single_double() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Using ms As New MemoryStream()
            Using bw As New BinaryWriter(ms, System.Text.Encoding.UTF8, True)
                bw.Write(1.5F)
                bw.Write(2.718281828459)
            End Using

            ms.Position = 0
            Using br As New BinaryReader(ms)
                Dim s = br.ReadSingle()
                Dim d = br.ReadDouble()
                Console.WriteLine(s & "|" & d)
            End Using
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1.5|2.718281828459"]);
}

#[test]
fn test_vb_binary_writer_reader_struct_serialization() {
    let src = r#"
Imports System.IO

Structure PlayerState
    Public ID As Integer
    Public Name As String
    Public Score As Double
End Structure

Module Program
    Sub Main()
        Dim player As New PlayerState With {.ID = 1, .Name = "Hero", .Score = 95.5}
        Using ms As New MemoryStream()
            Using bw As New BinaryWriter(ms, System.Text.Encoding.UTF8, True)
                bw.Write(player.ID)
                bw.Write(player.Name)
                bw.Write(player.Score)
            End Using

            ms.Position = 0
            Using br As New BinaryReader(ms)
                Dim restored As New PlayerState With {
                    .ID = br.ReadInt32(),
                    .Name = br.ReadString(),
                    .Score = br.ReadDouble()
                }
                Console.WriteLine(restored.ID & ":" & restored.Name & "=" & restored.Score)
            End Using
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1:Hero=95.5"]);
}

#[test]
fn test_vb_binary_writer_reader_enum_type() {
    let src = r#"
Imports System.IO

Enum GameState
    Menu = 1
    Playing = 2
End Enum

Module Program
    Sub Main()
        Using ms As New MemoryStream()
            Using bw As New BinaryWriter(ms, System.Text.Encoding.UTF8, True)
                bw.Write(CInt(GameState.Playing))
            End Using

            ms.Position = 0
            Using br As New BinaryReader(ms)
                Dim state As GameState = CType(br.ReadInt32(), GameState)
                Console.WriteLine(state.ToString())
            End Using
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Playing"]);
}

#[test]
fn test_vb_binary_writer_flush() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Using ms As New MemoryStream()
            Using bw As New BinaryWriter(ms, System.Text.Encoding.UTF8, True)
                bw.Write(777)
                bw.Flush()
                Console.WriteLine(ms.Length > 0)
            End Using
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_binary_reader_read_span_buffer() {
    let src = r#"
Imports System
Imports System.IO

Module Program
    Sub Main()
        Using ms As New MemoryStream()
            Using bw As New BinaryWriter(ms, System.Text.Encoding.UTF8, True)
                bw.Write(New Byte() {1, 2, 3, 4, 5})
            End Using

            ms.Position = 0
            Using br As New BinaryReader(ms)
                Dim buffer(4) As Byte
                Dim bytesRead = br.Read(buffer, 0, 5)
                Console.WriteLine(bytesRead & "|" & String.Join(",", buffer))
            End Using
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["5|1,2,3,4,5"]);
}

#[test]
fn test_vb_binary_writer_reader_sbyte_int16() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Using ms As New MemoryStream()
            Using bw As New BinaryWriter(ms, System.Text.Encoding.UTF8, True)
                bw.Write(CSByte(-50))
                bw.Write(CShort(-30000))
            End Using

            ms.Position = 0
            Using br As New BinaryReader(ms)
                Dim sb = br.ReadSByte()
                Dim s = br.ReadInt16()
                Console.WriteLine(sb & "|" & s)
            End Using
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["-50|-30000"]);
}

#[test]
fn test_vb_binary_writer_reader_unicode_encoding() {
    let src = r#"
Imports System.IO
Imports System.Text

Module Program
    Sub Main()
        Using ms As New MemoryStream()
            Using bw As New BinaryWriter(ms, Encoding.Unicode, True)
                bw.Write("UnicodeText")
            End Using

            ms.Position = 0
            Using br As New BinaryReader(ms, Encoding.Unicode)
                Dim str = br.ReadString()
                Console.WriteLine(str)
            End Using
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["UnicodeText"]);
}

#[test]
fn test_vb_binary_reader_base_stream_property() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Using ms As New MemoryStream()
            Using br As New BinaryReader(ms)
                Console.WriteLine(Object.ReferenceEquals(br.BaseStream, ms))
            End Using
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_binary_writer_base_stream_property() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Using ms As New MemoryStream()
            Using bw As New BinaryWriter(ms)
                Console.WriteLine(Object.ReferenceEquals(bw.BaseStream, ms))
            End Using
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_binary_writer_reader_multiple_objects_sequence() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Using ms As New MemoryStream()
            Using bw As New BinaryWriter(ms, System.Text.Encoding.UTF8, True)
                For i As Integer = 1 To 5
                    bw.Write("Item_" & i)
                Next
            End Using

            ms.Position = 0
            Using br As New BinaryReader(ms)
                For i As Integer = 1 To 5
                    Console.WriteLine(br.ReadString())
                Next
            End Using
        End Using
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["Item_1", "Item_2", "Item_3", "Item_4", "Item_5"]
    );
}

#[test]
fn test_vb_binary_writer_reader_empty_string() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Using ms As New MemoryStream()
            Using bw As New BinaryWriter(ms, System.Text.Encoding.UTF8, True)
                bw.Write("")
            End Using

            ms.Position = 0
            Using br As New BinaryReader(ms)
                Dim str = br.ReadString()
                Console.WriteLine(str.Length & "|" & (str = ""))
            End Using
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0|True"]);
}
