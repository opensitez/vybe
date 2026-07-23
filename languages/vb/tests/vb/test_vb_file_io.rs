use super::helpers::run_vb;

#[test]
fn file_write_all_text_and_read_all_text_roundtrip() {
    let out = run_vb(
        r#"
Imports System.IO

Module M
    Sub Main()
        Dim path As String = Path.GetTempFileName()
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
fn file_write_all_lines_and_read_all_lines_roundtrip() {
    let out = run_vb(
        r#"
Imports System.IO

Module M
    Sub Main()
        Dim path As String = Path.GetTempFileName()
        File.WriteAllLines(path, New String() {"a", "b", "c"})
        Dim lines As String() = File.ReadAllLines(path)
        Console.WriteLine(lines.Length)
        Console.WriteLine(lines(1))
        File.Delete(path)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["3", "b"]);
}

#[test]
fn file_append_all_text_accumulates() {
    let out = run_vb(
        r#"
Imports System.IO

Module M
    Sub Main()
        Dim path As String = Path.GetTempFileName()
        File.WriteAllText(path, "hello")
        File.AppendAllText(path, " world")
        Console.WriteLine(File.ReadAllText(path))
        File.Delete(path)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["hello world"]);
}

#[test]
fn file_exists_false_for_missing_path() {
    let out = run_vb(
        r#"
Imports System.IO

Module M
    Sub Main()
        Console.WriteLine(File.Exists("/no/such/file/xyz12345.txt"))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["False"]);
}

#[test]
fn file_copy_preserves_content() {
    let out = run_vb(
        r#"
Imports System.IO

Module M
    Sub Main()
        Dim source As String = Path.GetTempFileName()
        Dim target As String = source & ".copy"
        File.WriteAllText(source, "data")
        File.Copy(source, target, True)
        Console.WriteLine(File.ReadAllText(target))
        File.Delete(source)
        File.Delete(target)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["data"]);
}

#[test]
fn file_read_all_bytes_returns_expected_length() {
    let out = run_vb(
        r#"
Imports System.IO

Module M
    Sub Main()
        Dim path As String = Path.GetTempFileName()
        File.WriteAllBytes(path, New Byte() {1, 2, 3, 4, 5})
        Console.WriteLine(File.ReadAllBytes(path).Length)
        File.Delete(path)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["5"]);
}

#[test]
fn file_move_changes_location() {
    let out = run_vb(
        r#"
Imports System.IO

Module M
    Sub Main()
        Dim source As String = Path.GetTempFileName()
        Dim target As String = source & ".moved"
        File.WriteAllText(source, "x")
        File.Move(source, target)
        Console.WriteLine(File.Exists(source))
        Console.WriteLine(File.ReadAllText(target))
        File.Delete(target)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["False", "x"]);
}

#[test]
fn file_append_text_then_read_lines_count() {
    let out = run_vb(
        r#"
Imports System.IO

Module M
    Sub Main()
        Dim path As String = Path.GetTempFileName()
        File.WriteAllText(path, "one\n")
        File.AppendAllText(path, "two\n")
        Dim lines As String() = File.ReadAllText(path).Split("\n"c)
        Console.WriteLine(lines.Length)
        Console.WriteLine(lines(0))
        File.Delete(path)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["3", "one"]);
}
