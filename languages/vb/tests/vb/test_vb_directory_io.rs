use super::helpers::run_vb;

#[test]
fn directory_create_exists_and_delete() {
    let out = run_vb(
        r#"
Imports System.IO

Module M
    Sub Main()
        Dim path As String = Path.Combine(Path.GetTempPath(), "vybe_dir_" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(path)
        Console.WriteLine(Directory.Exists(path))
        Directory.Delete(path)
        Console.WriteLine(Directory.Exists(path))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn directory_exists_false_for_missing_path() {
    let out = run_vb(
        r#"
Imports System.IO

Module M
    Sub Main()
        Console.WriteLine(Directory.Exists("/no/such/path/xyz12345"))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["False"]);
}

#[test]
fn directory_get_files_returns_created_entries_count() {
    let out = run_vb(
        r#"
Imports System.IO

Module M
    Sub Main()
        Dim dir As String = Path.Combine(Path.GetTempPath(), "vybe_files_" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(dir)

        File.WriteAllText(Path.Combine(dir, "a.txt"), "1")
        File.WriteAllText(Path.Combine(dir, "b.txt"), "2")
        File.WriteAllText(Path.Combine(dir, "c.log"), "3")

        Console.WriteLine(Directory.GetFiles(dir).Length)
        Console.WriteLine(Directory.GetFiles(dir, "*.txt").Length)

        Directory.Delete(dir, True)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["3", "2"]);
}

#[test]
fn directory_get_directories_counts_items() {
    let out = run_vb(
        r#"
Imports System.IO

Module M
    Sub Main()
        Dim root As String = Path.Combine(Path.GetTempPath(), "vybe_tree_" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(root)
        Directory.CreateDirectory(Path.Combine(root, "a"))
        Directory.CreateDirectory(Path.Combine(root, "b"))

        Console.WriteLine(Directory.GetDirectories(root).Length)

        Directory.Delete(root, True)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["2"]);
}

#[test]
fn directory_move_directory_renames() {
    let out = run_vb(
        r#"
Imports System.IO

Module M
    Sub Main()
        Dim source As String = Path.Combine(Path.GetTempPath(), "vybe_src_" & Guid.NewGuid().ToString("N"))
        Dim target As String = Path.Combine(Path.GetTempPath(), "vybe_dst_" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(source)
        Directory.Move(source, target)

        Console.WriteLine(Directory.Exists(source))
        Console.WriteLine(Directory.Exists(target))

        Directory.Delete(target)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["False", "True"]);
}

#[test]
fn directory_enumeration_handles_nested_dirs() {
    let out = run_vb(
        r#"
Imports System.IO

Module M
    Sub Main()
        Dim root As String = Path.Combine(Path.GetTempPath(), "vybe_nested_" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(root)
        Directory.CreateDirectory(Path.Combine(root, "a"))
        Directory.CreateDirectory(Path.Combine(root, "a", "b"))

        Dim topLevel As String() = Directory.GetDirectories(root)
        Dim nested As String() = Directory.GetDirectories(Path.Combine(root, "a"))
        Console.WriteLine(topLevel.Length)
        Console.WriteLine(nested.Length)

        Directory.Delete(root, True)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["1", "1"]);
}
