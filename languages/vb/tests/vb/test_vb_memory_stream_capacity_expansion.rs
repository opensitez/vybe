use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: MemoryStream Capacity, Resizing & Buffer Operations
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_memory_stream_initial_capacity() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Using ms As New MemoryStream(1024)
            Console.WriteLine(ms.Capacity & "|" & ms.Length)
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1024|0"]);
}

#[test]
fn test_vb_memory_stream_auto_capacity_expansion() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Using ms As New MemoryStream(2)
            Dim initialCap = ms.Capacity
            ms.Write({1, 2, 3, 4, 5}, 0, 5)
            Console.WriteLine(initialCap < ms.Capacity & "|" & ms.Length)
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|5"]);
}

#[test]
fn test_vb_memory_stream_get_buffer_expose_internal_array() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Using ms As New MemoryStream()
            ms.WriteByte(100)
            ms.WriteByte(200)
            Dim buffer = ms.GetBuffer()
            Console.WriteLine(buffer(0) & "|" & buffer(1))
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["100|200"]);
}

#[test]
fn test_vb_memory_stream_try_get_buffer_segment() {
    let src = r#"
Imports System
Imports System.IO

Module Program
    Sub Main()
        Using ms As New MemoryStream()
            ms.WriteByte(50)
            Dim segment As ArraySegment(Of Byte)
            Dim ok = ms.TryGetBuffer(segment)
            Console.WriteLine(ok & ":" & segment.Array(0))
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True:50"]);
}

#[test]
fn test_vb_memory_stream_to_array_creates_exact_copy() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Using ms As New MemoryStream(100)
            ms.Write({10, 20, 30}, 0, 3)
            Dim arr = ms.ToArray()
            Console.WriteLine(arr.Length & "|" & ms.Capacity)
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3|100"]);
}

#[test]
fn test_vb_memory_stream_set_length_truncate() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Using ms As New MemoryStream()
            ms.Write({1, 2, 3, 4, 5}, 0, 5)
            ms.SetLength(3)
            Console.WriteLine(ms.Length & "|" & String.Join(",", ms.ToArray()))
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3|1,2,3"]);
}

#[test]
fn test_vb_memory_stream_set_length_expand() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Using ms As New MemoryStream()
            ms.WriteByte(42)
            ms.SetLength(5)
            Console.WriteLine(ms.Length & "|" & String.Join(",", ms.ToArray()))
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["5|42,0,0,0,0"]);
}

#[test]
fn test_vb_memory_stream_seek_begin_current_end() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Using ms As New MemoryStream()
            ms.Write({10, 20, 30, 40, 50}, 0, 5)
            ms.Seek(1, SeekOrigin.Begin)
            Dim b1 = ms.ReadByte()
            ms.Seek(-1, SeekOrigin.End)
            Dim b2 = ms.ReadByte()
            Console.WriteLine(b1 & "|" & b2)
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["20|50"]);
}

#[test]
fn test_vb_memory_stream_position_property_get_set() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Using ms As New MemoryStream()
            ms.WriteByte(99)
            Console.WriteLine("Pos: " & ms.Position)
            ms.Position = 0
            Console.WriteLine("Val: " & ms.ReadByte())
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Pos: 1", "Val: 99"]);
}

#[test]
fn test_vb_memory_stream_fixed_buffer_constructor_cannot_expand() {
    let src = r#"
Imports System
Imports System.IO

Module Program
    Sub Main()
        Dim fixedBuffer As Byte() = New Byte(4) {}
        Using ms As New MemoryStream(fixedBuffer)
            Try
                ms.SetLength(10)
            Catch ex As NotSupportedException
                Console.WriteLine("NotSupportedException Caught on Fixed MemoryStream")
            End Try
        End Using
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["NotSupportedException Caught on Fixed MemoryStream"]
    );
}

#[test]
fn test_vb_memory_stream_fixed_buffer_write_past_capacity_throws() {
    let src = r#"
Imports System
Imports System.IO

Module Program
    Sub Main()
        Dim fixedBuffer As Byte() = New Byte(1) {} ' Length 2
        Using ms As New MemoryStream(fixedBuffer)
            Try
                ms.Write({1, 2, 3, 4}, 0, 4)
            Catch ex As NotSupportedException
                Console.WriteLine("NotSupportedException Caught on Write Overflow")
            End Try
        End Using
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["NotSupportedException Caught on Write Overflow"]
    );
}

#[test]
fn test_vb_memory_stream_capacity_property_manual_shrink() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Using ms As New MemoryStream(100)
            ms.WriteByte(1)
            ms.Capacity = 10
            Console.WriteLine(ms.Capacity & "|" & ms.Length)
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10|1"]);
}

#[test]
fn test_vb_memory_stream_capacity_smaller_than_length_throws() {
    let src = r#"
Imports System
Imports System.IO

Module Program
    Sub Main()
        Using ms As New MemoryStream()
            ms.Write({1, 2, 3, 4, 5}, 0, 5)
            Try
                ms.Capacity = 2
            Catch ex As ArgumentOutOfRangeException
                Console.WriteLine("ArgumentOutOfRangeException Caught")
            End Try
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["ArgumentOutOfRangeException Caught"]);
}

#[test]
fn test_vb_memory_stream_get_buffer_throws_when_not_publicly_visible() {
    let src = r#"
Imports System
Imports System.IO

Module Program
    Sub Main()
        Dim data As Byte() = {1, 2, 3}
        ' MemoryStream constructed with publiclyVisible=false
        Using ms As New MemoryStream(data, 0, 3, False, False)
            Try
                ms.GetBuffer()
            Catch ex As UnauthorizedAccessException
                Console.WriteLine("UnauthorizedAccessException Caught")
            End Try
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["UnauthorizedAccessException Caught"]);
}

#[test]
fn test_vb_memory_stream_can_seek_can_read_can_write() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Using ms As New MemoryStream()
            Console.WriteLine(ms.CanSeek & "|" & ms.CanRead & "|" & ms.CanWrite)
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True|True"]);
}

#[test]
fn test_vb_memory_stream_write_span_byte() {
    let src = r#"
Imports System
Imports System.IO

Module Program
    Sub Main()
        Using ms As New MemoryStream()
            Dim span As ReadOnlySpan(Of Byte) = New Byte() {1, 2, 3}
            ms.Write(span)
            Console.WriteLine(ms.Length)
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3"]);
}

#[test]
fn test_vb_memory_stream_read_span_byte() {
    let src = r#"
Imports System
Imports System.IO

Module Program
    Sub Main()
        Using ms As New MemoryStream(New Byte() {10, 20, 30})
            Dim buffer(2) As Byte
            Dim span As Span(Of Byte) = buffer
            Dim readCount = ms.Read(span)
            Console.WriteLine(readCount & ":" & String.Join(",", buffer))
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3:10,20,30"]);
}

#[test]
fn test_vb_memory_stream_closed_stream_throws_object_disposed() {
    let src = r#"
Imports System
Imports System.IO

Module Program
    Sub Main()
        Dim ms As New MemoryStream()
        ms.Close()
        Try
            ms.WriteByte(1)
        Catch ex As ObjectDisposedException
            Console.WriteLine("ObjectDisposedException Caught")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["ObjectDisposedException Caught"]);
}

#[test]
fn test_vb_memory_stream_write_to_another_stream() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Using ms1 As New MemoryStream()
            ms1.Write({10, 20, 30}, 0, 3)
            Using ms2 As New MemoryStream()
                ms1.WriteTo(ms2)
                Console.WriteLine(String.Join(",", ms2.ToArray()))
            End Using
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10,20,30"]);
}

#[test]
fn test_vb_memory_stream_construct_offset_count() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Dim raw As Byte() = {10, 20, 30, 40, 50}
        ' Sub-slice offset 1, length 3
        Using ms As New MemoryStream(raw, 1, 3)
            Console.WriteLine(ms.Length & "|" & String.Join(",", ms.ToArray()))
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3|20,30,40"]);
}
