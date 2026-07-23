use super::helpers::run_vb;

#[test]
fn path_combine_merges_segments() {
    let out = run_vb(
        r#"
Imports System.IO

Module M
    Sub Main()
        Console.WriteLine(Path.Combine("home", "user", "doc.txt"))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["home/user/doc.txt"]);
}

#[test]
fn path_get_temp_path_is_non_empty() {
    let out = run_vb(
        r#"
Imports System.IO

Module M
    Sub Main()
        Console.WriteLine(Path.GetTempPath().Length > 0)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}

#[test]
fn path_get_file_name_and_extension() {
    let out = run_vb(
        r#"
Imports System.IO

Module M
    Sub Main()
        Console.WriteLine(Path.GetFileName("/tmp/archive.tar.gz"))
        Console.WriteLine(Path.GetExtension("/tmp/archive.tar.gz"))
        Console.WriteLine(Path.GetFileNameWithoutExtension("/tmp/archive.tar.gz"))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["archive.tar.gz", ".gz", "archive.tar"]);
}

#[test]
fn path_change_extension_preserves_name() {
    let out = run_vb(
        r#"
Imports System.IO

Module M
    Sub Main()
        Console.WriteLine(Path.ChangeExtension("data.bin", ".txt"))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["data.txt"]);
}

#[test]
fn path_get_directory_name_returns_parent() {
    let out = run_vb(
        r#"
Imports System.IO

Module M
    Sub Main()
        Console.WriteLine(Path.GetDirectoryName("/tmp/logs/app.log"))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["/tmp/logs"]);
}

#[test]
fn path_get_full_path_normalizes_relative() {
    let out = run_vb(
        r#"
Imports System.IO

Module M
    Sub Main()
        Console.WriteLine(Path.GetFullPath("./tmp").Length > 0)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}

#[test]
fn path_has_extension_check() {
    let out = run_vb(
        r#"
Imports System.IO

Module M
    Sub Main()
        Console.WriteLine(Path.HasExtension("/tmp/file.txt"))
        Console.WriteLine(Path.HasExtension("/tmp/file"))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn path_get_temp_file_name_extension() {
    let out = run_vb(
        r#"
Imports System.IO

Module M
    Sub Main()
        Dim p As String = Path.GetTempFileName()
        Console.WriteLine(Path.GetExtension(p))
        Console.WriteLine(p.Contains(System.IO.Path.GetTempPath()))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec![".tmp", "True"]);
}

#[test]
fn path_get_path_root_reports_root_segment() {
    let out = run_vb(
        r#"
Imports System.IO

Module M
    Sub Main()
        Console.WriteLine(Path.GetPathRoot("/tmp/file.txt") = "/")
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}
