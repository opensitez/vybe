use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: System.IO.StringWriter & StringReader Mechanics
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_string_writer_write_line() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Using sw As New StringWriter()
            sw.WriteLine("Line 1")
            sw.WriteLine("Line 2")
            Console.WriteLine(sw.ToString().Contains("Line 1") AndAlso sw.ToString().Contains("Line 2"))
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_string_reader_read_line_sequence() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Dim text = "First Line" & vbCrLf & "Second Line"
        Using sr As New StringReader(text)
            Dim l1 = sr.ReadLine()
            Dim l2 = sr.ReadLine()
            Console.WriteLine(l1 & "|" & l2)
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["First Line|Second Line"]);
}

#[test]
fn test_vb_string_reader_read_to_end() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Dim text = "Header" & vbLf & "Body" & vbLf & "Footer"
        Using sr As New StringReader(text)
            Dim line1 = sr.ReadLine()
            Dim rest = sr.ReadToEnd()
            Console.WriteLine(line1 & "|" & rest.Replace(vbLf, "-"))
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Header|Body-Footer"]);
}

#[test]
fn test_vb_string_reader_peek_character() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Using sr As New StringReader("ABC")
            Dim p1 = sr.Peek()
            Dim ch1 = sr.Read()
            Dim p2 = sr.Peek()
            Console.WriteLine(ChrW(p1) & "=" & ChrW(ch1) & "|Next=" & ChrW(p2))
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["A=A|Next=B"]);
}

#[test]
fn test_vb_string_reader_read_buffer_array() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Using sr As New StringReader("VisualBasic")
            Dim buffer(5) As Char
            Dim count = sr.Read(buffer, 0, 6)
            Console.WriteLine(count & ":" & New String(buffer))
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["6:Visual"]);
}

#[test]
fn test_vb_string_writer_get_string_builder() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Using sw As New StringWriter()
            sw.Write("Hello ")
            Dim sb = sw.GetStringBuilder()
            sb.Append("World")
            Console.WriteLine(sw.ToString())
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Hello World"]);
}

#[test]
fn test_vb_string_writer_custom_new_line() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Using sw As New StringWriter()
            sw.NewLine = ";"
            sw.WriteLine("A")
            sw.WriteLine("B")
            Console.WriteLine(sw.ToString())
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["A;B;"]);
}

#[test]
fn test_vb_string_writer_write_formatted_arguments() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Using sw As New StringWriter()
            sw.Write("Item {0} cost {1:C}", "Widget", 19.99)
            Console.WriteLine(sw.ToString().Contains("Widget"))
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_string_writer_encoding_property() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Using sw As New StringWriter()
            Console.WriteLine(sw.Encoding.WebName)
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["utf-16"]);
}

#[test]
fn test_vb_string_reader_empty_string_read_line_returns_null() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Using sr As New StringReader("")
            Dim line = sr.ReadLine()
            Console.WriteLine(line Is Nothing)
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_string_reader_read_end_of_stream_returns_minus_one() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Using sr As New StringReader("A")
            sr.Read()
            Dim ch = sr.Read()
            Dim peek = sr.Peek()
            Console.WriteLine(ch & "|" & peek)
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["-1|-1"]);
}

#[test]
fn test_vb_string_writer_write_char_array() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Using sw As New StringWriter()
            Dim chars As Char() = {"X"c, "Y"c, "Z"c}
            sw.Write(chars, 1, 2)
            Console.WriteLine(sw.ToString())
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["YZ"]);
}

#[test]
fn test_vb_string_writer_write_primitives() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Using sw As New StringWriter()
            sw.Write(10)
            sw.Write(True)
            sw.Write(3.14)
            Console.WriteLine(sw.ToString())
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10True3.14"]);
}

#[test]
fn test_vb_string_reader_read_async_simulation() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Dim text = "AsyncLine1" & vbCrLf & "AsyncLine2"
        Using sr As New StringReader(text)
            Dim t1 = sr.ReadLineAsync()
            Dim t2 = sr.ReadLineAsync()
            Console.WriteLine(t1.Result & "|" & t2.Result)
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["AsyncLine1|AsyncLine2"]);
}

#[test]
fn test_vb_string_writer_write_line_async_simulation() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Using sw As New StringWriter()
            Dim t1 = sw.WriteLineAsync("AsyncWritten")
            t1.Wait()
            Console.WriteLine(sw.ToString().Trim())
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["AsyncWritten"]);
}

#[test]
fn test_vb_string_writer_with_custom_string_builder() {
    let src = r#"
Imports System.IO
Imports System.Text

Module Program
    Sub Main()
        Dim sb As New StringBuilder("Initial: ")
        Using sw As New StringWriter(sb)
            sw.Write("Appended")
        End Using
        Console.WriteLine(sb.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Initial: Appended"]);
}

#[test]
fn test_vb_string_reader_read_block() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Using sr As New StringReader("0123456789")
            Dim buffer(4) As Char
            Dim readCount = sr.ReadBlock(buffer, 0, 5)
            Console.WriteLine(readCount & ":" & New String(buffer))
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["5:01234"]);
}

#[test]
fn test_vb_string_writer_flush_does_not_clear() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Using sw As New StringWriter()
            sw.Write("Data")
            sw.Flush()
            Console.WriteLine(sw.ToString())
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Data"]);
}

#[test]
fn test_vb_string_writer_close_and_dispose() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Dim sw As New StringWriter()
        sw.Write("Content")
        sw.Close()
        ' StringWriter.ToString still returns content after Dispose/Close!
        Console.WriteLine(sw.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Content"]);
}

#[test]
fn test_vb_string_reader_close_and_read_throws() {
    let src = r#"
Imports System
Imports System.IO

Module Program
    Sub Main()
        Dim sr As New StringReader("Text")
        sr.Close()
        Try
            Dim line = sr.ReadLine()
        Catch ex As ObjectDisposedException
            Console.WriteLine("ObjectDisposedException Caught")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["ObjectDisposedException Caught"]);
}
