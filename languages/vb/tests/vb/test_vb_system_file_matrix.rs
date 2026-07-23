use super::helpers::run_vb;

#[test]
fn file_write_and_read_text_roundtrip() {
    let out = run_vb(
        r#"
Imports System
Imports System.IO

Module M
    Sub Main()
        Dim path As String = Path.Combine(Path.GetTempPath(), "vb_file_text_" & Guid.NewGuid().ToString("N"))
        File.WriteAllText(path, "hello")
        Console.WriteLine(File.ReadAllText(path))
        File.Delete(path)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["hello"]);
}

#[test]
fn file_append_text_accumulates() {
    let out = run_vb(
        r#"
Imports System
Imports System.IO

Module M
    Sub Main()
        Dim path As String = Path.Combine(Path.GetTempPath(), "vb_file_append_" & Guid.NewGuid().ToString("N"))
        File.WriteAllText(path, "left")
        File.AppendAllText(path, "-right")
        Console.WriteLine(File.ReadAllText(path))
        File.Delete(path)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["left-right"]);
}

#[test]
fn file_bytes_roundtrip() {
    let out = run_vb(
        r#"
Imports System
Imports System.IO

Module M
    Sub Main()
        Dim path As String = Path.Combine(Path.GetTempPath(), "vb_file_bytes_" & Guid.NewGuid().ToString("N"))
        Dim input() As Byte = {1, 2, 3, 4, 5}
        File.WriteAllBytes(path, input)
        Dim output() As Byte = File.ReadAllBytes(path)
        Console.WriteLine(output.Length)
        Console.WriteLine(output(0))
        Console.WriteLine(output(4))
        File.Delete(path)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["5", "1", "5"]);
}

#[test]
fn file_copy_duplicated_content() {
    let out = run_vb(
        r#"
Imports System.IO

Module M
    Sub Main()
        Dim source As String = Path.Combine(Path.GetTempPath(), "vb_file_src_" & Guid.NewGuid().ToString("N"))
        Dim destination As String = Path.Combine(Path.GetTempPath(), "vb_file_dst_" & Guid.NewGuid().ToString("N"))
        File.WriteAllText(source, "copy-me")
        File.Copy(source, destination, True)
        Console.WriteLine(File.ReadAllText(destination))
        Console.WriteLine(File.ReadAllText(source) = File.ReadAllText(destination))
        File.Delete(source)
        File.Delete(destination)
    End Module
End Module
"#,
    );

    assert_eq!(out, vec!["copy-me", "True"]);
}

#[test]
fn file_move_changes_location() {
    let out = run_vb(
        r#"
Imports System.IO

Module M
    Sub Main()
        Dim source As String = Path.Combine(Path.GetTempPath(), "vb_file_move_src_" & Guid.NewGuid().ToString("N"))
        Dim destination As String = Path.Combine(Path.GetTempPath(), "vb_file_move_dst_" & Guid.NewGuid().ToString("N"))
        File.WriteAllText(source, "m")
        File.Move(source, destination)
        Console.WriteLine(File.Exists(source))
        Console.WriteLine(File.Exists(destination))
        Console.WriteLine(File.ReadAllText(destination))
        File.Delete(destination)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["False", "True", "m"]);
}

#[test]
fn file_exists_is_false_for_missing_file() {
    let out = run_vb(
        r#"
Imports System.IO

Module M
    Sub Main()
        Dim path As String = Path.Combine(Path.GetTempPath(), "vb_file_missing_" & Guid.NewGuid().ToString("N"))
        Console.WriteLine(File.Exists(path))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["False"]);
}

#[test]
fn file_lines_reading() {
    let out = run_vb(
        r#"
Imports System.IO

Module M
    Sub Main()
        Dim path As String = Path.Combine(Path.GetTempPath(), "vb_file_lines_" & Guid.NewGuid().ToString("N"))
        File.WriteAllText(path, "a" & vbLf & "b" & vbLf & "c")
        Dim lines() As String = File.ReadAllLines(path)
        Console.WriteLine(lines.Length)
        Console.WriteLine(lines(0))
        Console.WriteLine(lines(2))
        File.Delete(path)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["3", "a", "c"]);
}

#[test]
fn file_attributes_and_extension() {
    let out = run_vb(
        r#"
Imports System.IO

Module M
    Sub Main()
        Dim path As String = Path.Combine(Path.GetTempPath(), "vb_ext." & Guid.NewGuid().ToString("N") & ".bin")
        File.WriteAllText(path, "x")
        Dim info As New FileInfo(path)
        Console.WriteLine(info.Extension = ".bin")
        Console.WriteLine(info.Exists)
        File.Delete(path)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn file_create_with_temporary_handle() {
    let out = run_vb(
        r#"
Imports System.IO

Module M
    Sub Main()
        Dim path As String = Path.Combine(Path.GetTempPath(), "vb_create_" & Guid.NewGuid().ToString("N"))
        Using fs As FileStream = File.Create(path)
            fs.WriteByte(10)
        End Using
        Using fs2 As FileStream = File.OpenRead(path)
            Console.WriteLine(fs2.Length)
        End Using
        File.Delete(path)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["1"]);
}
