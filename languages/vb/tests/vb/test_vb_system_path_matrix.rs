use super::helpers::run_vb;

#[test]
fn path_change_extension_replaces_extension() {
    let out = run_vb(
        r#"
Imports System.IO

Module M
    Sub Main()
        Console.WriteLine(Path.ChangeExtension("report.txt", ".json"))
        Console.WriteLine(Path.ChangeExtension("archive", ".bin"))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["report.json", "archive.bin"]);
}

#[test]
fn path_get_full_path_and_root() {
    let out = run_vb(
        r#"
Imports System.IO

Module M
    Sub Main()
        Dim absolute As String = Path.GetFullPath("./")
        Console.WriteLine(absolute.Length > 0)
        Console.WriteLine(Path.GetPathRoot(absolute).Length > 0)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn path_get_random_file_name_is_valid_pattern() {
    let out = run_vb(
        r#"
Imports System.IO

Module M
    Sub Main()
        Dim randomName As String = Path.GetRandomFileName()
        Console.WriteLine(randomName.Length >= 8)
        Console.WriteLine(Path.GetFileName(randomName) = randomName)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn path_get_temp_file_name_has_tmp_extension() {
    let out = run_vb(
        r#"
Imports System.IO

Module M
    Sub Main()
        Dim temp As String = Path.GetTempFileName()
        Console.WriteLine(Path.GetExtension(temp))
        Console.WriteLine(Path.GetDirectoryName(temp) = Path.GetTempPath())
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec![".tmp", "True"]);
}

#[test]
fn path_relative_to_root_and_has_extension() {
    let out = run_vb(
        r#"
Imports System.IO

Module M
    Sub Main()
        Dim p As String = Path.Combine("app", "logs", "trace.log")
        Console.WriteLine(Path.HasExtension(p))
        Console.WriteLine(Path.GetFileNameWithoutExtension(p))
        Console.WriteLine(Path.GetDirectoryName(p).Contains("app" & Path.DirectorySeparatorChar & "logs"))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "trace", "app\\logs"]);
}

#[test]
fn path_get_file_name_and_root() {
    let out = run_vb(
        r#"
Imports System.IO

Module M
    Sub Main()
        Dim p As String = "/tmp/reports/final.vb"
        Console.WriteLine(Path.GetFileName(p))
        Console.WriteLine(Path.GetExtension(p))
        Console.WriteLine(Path.GetDirectoryName(p))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["final.vb", ".vb", "/tmp/reports"]);
}

#[test]
fn path_get_invalid_file_name_chars_is_present_set() {
    let out = run_vb(
        r#"
Imports System.IO

Module M
    Sub Main()
        Dim bad() As Char = Path.GetInvalidFileNameChars()
        Dim badCount As Integer = bad.Length

        Dim ok As Boolean = False
        For i As Integer = 0 To bad.Length - 1
            If bad(i) <> ""C Then
                ok = True
                Exit For
            End If
        Next

        Console.WriteLine(badCount >= 0)
        Console.WriteLine(ok)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn path_is_path_rooted_and_get_filename() {
    let out = run_vb(
        r#"
Imports System.IO

Module M
    Sub Main()
        Console.WriteLine(Path.IsPathRooted("/tmp/a.txt"))
        Console.WriteLine(Path.IsPathRooted("a/b.txt"))
        Console.WriteLine(Path.GetFileName("C:\\Users\\me\\data.bin"))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "False", "data.bin"]);
}

#[test]
fn path_combine_nested_segments() {
    let out = run_vb(
        r#"
Imports System.IO

Module M
    Sub Main()
        Dim combined As String = Path.Combine("/tmp", "a", "b", "c.txt")
        Console.WriteLine(combined)
        Console.WriteLine(Path.GetFileName(combined))
        Console.WriteLine(Path.GetDirectoryName(combined))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["/tmp/a/b/c.txt", "c.txt", "/tmp/a/b"]);
}

#[test]
fn path_get_temppath_is_stable() {
    let out = run_vb(
        r#"
Imports System.IO

Module M
    Sub Main()
        Dim a As String = Path.GetTempPath()
        Dim b As String = Path.GetTempPath()
        Console.WriteLine(a.Length > 0)
        Console.WriteLine(a = b)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}
