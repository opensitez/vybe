use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Stream.CopyTo & Stream.CopyToAsync Operations
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_stream_copy_to_synchronous() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Dim sourceData As Byte() = {1, 2, 3, 4, 5}
        Using srcMs As New MemoryStream(sourceData)
            Using destMs As New MemoryStream()
                srcMs.CopyTo(destMs)
                Console.WriteLine(destMs.Length & "|" & String.Join(",", destMs.ToArray()))
            End Using
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["5|1,2,3,4,5"]);
}

#[test]
fn test_vb_stream_copy_to_async_basic() {
    let src = r#"
Imports System.IO
Imports System.Threading.Tasks

Module Program
    Private Async Function CopyStreamAsync() As Task(Of Integer)
        Dim data As Byte() = {10, 20, 30}
        Using srcMs As New MemoryStream(data)
            Using destMs As New MemoryStream()
                Await srcMs.CopyToAsync(destMs)
                Return CInt(destMs.Length)
            End Using
        End Using
    End Function

    Sub Main()
        Dim t = CopyStreamAsync()
        Console.WriteLine(t.Result)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3"]);
}

#[test]
fn test_vb_stream_copy_to_with_custom_buffer_size() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Dim data As Byte() = New Byte(99) {}
        For i As Integer = 0 To 99 : data(i) = CByte(i) : Next

        Using srcMs As New MemoryStream(data)
            Using destMs As New MemoryStream()
                srcMs.CopyTo(destMs, 16) ' 16-byte buffer size
                Console.WriteLine(destMs.Length)
            End Using
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["100"]);
}

#[test]
fn test_vb_stream_copy_to_async_cancellation_token() {
    let src = r#"
Imports System.IO
Imports System.Threading
Imports System.Threading.Tasks

Module Program
    Sub Main()
        Dim cts As New CancellationTokenSource()
        cts.Cancel()

        Dim data As Byte() = {1, 2, 3}
        Using srcMs As New MemoryStream(data)
            Using destMs As New MemoryStream()
                Dim t = srcMs.CopyToAsync(destMs, 8192, cts.Token)
                Console.WriteLine(t.IsCanceled)
            End Using
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_stream_copy_to_partial_source_position() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Dim data As Byte() = {10, 20, 30, 40, 50}
        Using srcMs As New MemoryStream(data)
            srcMs.Position = 2 ' Skip first 2 bytes
            Using destMs As New MemoryStream()
                srcMs.CopyTo(destMs)
                Console.WriteLine(String.Join(",", destMs.ToArray()))
            End Using
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["30,40,50"]);
}

#[test]
fn test_vb_stream_copy_to_append_dest_position() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Dim d1 As Byte() = {1, 2}
        Dim d2 As Byte() = {3, 4}
        Using destMs As New MemoryStream()
            Using s1 As New MemoryStream(d1)
                s1.CopyTo(destMs)
            End Using
            Using s2 As New MemoryStream(d2)
                s2.CopyTo(destMs)
            End Using
            Console.WriteLine(String.Join(",", destMs.ToArray()))
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1,2,3,4"]);
}

#[test]
fn test_vb_stream_copy_to_empty_source() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Using srcMs As New MemoryStream()
            Using destMs As New MemoryStream()
                srcMs.CopyTo(destMs)
                Console.WriteLine(destMs.Length)
            End Using
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0"]);
}

#[test]
fn test_vb_stream_copy_to_unreadable_source_throws() {
    let src = r#"
Imports System
Imports System.IO

Module Program
    Sub Main()
        ' MemoryStream constructed with writable=false and publiclyVisible=false is read-only, but let's test a closed stream!
        Dim srcMs As New MemoryStream({1, 2, 3})
        srcMs.Close()
        Using destMs As New MemoryStream()
            Try
                srcMs.CopyTo(destMs)
            Catch ex As ObjectDisposedException
                Console.WriteLine("ObjectDisposedException Caught")
            End Try
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["ObjectDisposedException Caught"]);
}

#[test]
fn test_vb_stream_copy_to_unwritable_destination_throws() {
    let src = r#"
Imports System
Imports System.IO

Module Program
    Sub Main()
        Dim readOnlyDest As New MemoryStream({0, 0, 0}, False) ' Writable = False
        Using srcMs As New MemoryStream({1, 2, 3})
            Try
                srcMs.CopyTo(readOnlyDest)
            Catch ex As NotSupportedException
                Console.WriteLine("NotSupportedException Caught on ReadOnly Dest")
            End Try
        End Using
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["NotSupportedException Caught on ReadOnly Dest"]
    );
}

#[test]
fn test_vb_stream_read_write_timeout_properties() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Using ms As New MemoryStream()
            Console.WriteLine(ms.CanTimeout)
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False"]);
}

#[test]
fn test_vb_stream_flush_async_simulation() {
    let src = r#"
Imports System.IO
Imports System.Threading.Tasks

Module Program
    Private Async Function FlushAsyncTest() As Task
        Using ms As New MemoryStream()
            ms.Write({10, 20}, 0, 2)
            Await ms.FlushAsync()
            Console.WriteLine("Flushed Async")
        End Using
    End Function

    Sub Main()
        Dim t = FlushAsyncTest()
        t.Wait()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Flushed Async"]);
}

#[test]
fn test_vb_stream_read_async_buffer() {
    let src = r#"
Imports System.IO
Imports System.Threading.Tasks

Module Program
    Private Async Function ReadAsyncTest() As Task(Of Integer)
        Dim data As Byte() = {100, 200}
        Using ms As New MemoryStream(data)
            Dim buffer(1) As Byte
            Dim readCount = Await ms.ReadAsync(buffer, 0, 2)
            Return readCount
        End Using
    End Function

    Sub Main()
        Dim t = ReadAsyncTest()
        Console.WriteLine(t.Result)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2"]);
}

#[test]
fn test_vb_stream_write_async_buffer() {
    let src = r#"
Imports System.IO
Imports System.Threading.Tasks

Module Program
    Private Async Function WriteAsyncTest() As Task(Of Long)
        Using ms As New MemoryStream()
            Dim buffer As Byte() = {50, 60, 70}
            Await ms.WriteAsync(buffer, 0, 3)
            Return ms.Length
        End Using
    End Function

    Sub Main()
        Dim t = WriteAsyncTest()
        Console.WriteLine(t.Result)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3"]);
}

#[test]
fn test_vb_stream_null_stream_singleton() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Dim nullStrm = Stream.Null
        Console.WriteLine(nullStrm.CanRead & "|" & nullStrm.CanWrite & "|" & nullStrm.Length)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True|0"]);
}

#[test]
fn test_vb_stream_copy_to_null_stream() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Dim data As Byte() = {1, 2, 3, 4, 5}
        Using srcMs As New MemoryStream(data)
            srcMs.CopyTo(Stream.Null)
            Console.WriteLine(srcMs.Position)
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["5"]);
}

#[test]
fn test_vb_stream_synchronized_wrapper() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Using ms As New MemoryStream()
            Dim syncStrm = Stream.Synchronized(ms)
            Console.WriteLine(syncStrm.CanWrite)
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_stream_dispose_async_simulation() {
    let src = r#"
Imports System.IO
Imports System.Threading.Tasks

Module Program
    Private Async Function DisposeAsyncTest() As Task
        Dim ms As New MemoryStream()
        ms.WriteByte(42)
        Await ms.DisposeAsync()
        Console.WriteLine("Disposed Async")
    End Function

    Sub Main()
        Dim t = DisposeAsyncTest()
        t.Wait()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Disposed Async"]);
}

#[test]
fn test_vb_stream_copy_to_invalid_buffer_size_throws() {
    let src = r#"
Imports System
Imports System.IO

Module Program
    Sub Main()
        Using srcMs As New MemoryStream({1})
            Using destMs As New MemoryStream()
                Try
                    srcMs.CopyTo(destMs, 0)
                Catch ex As ArgumentOutOfRangeException
                    Console.WriteLine("ArgumentOutOfRangeException Caught")
                End Try
            End Using
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["ArgumentOutOfRangeException Caught"]);
}

#[test]
fn test_vb_stream_copy_to_null_destination_throws() {
    let src = r#"
Imports System
Imports System.IO

Module Program
    Sub Main()
        Using srcMs As New MemoryStream({1})
            Try
                srcMs.CopyTo(Nothing)
            Catch ex As ArgumentNullException
                Console.WriteLine("ArgumentNullException Caught")
            End Try
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["ArgumentNullException Caught"]);
}

#[test]
fn test_vb_stream_read_byte_write_byte() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Using ms As New MemoryStream()
            ms.WriteByte(255)
            ms.Position = 0
            Dim b = ms.ReadByte()
            Console.WriteLine(b)
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["255"]);
}
