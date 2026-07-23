use super::helpers::run_vb;

#[test]
fn directory_create_and_exists() {
    let out = run_vb(
        r#"
Imports System
Imports System.IO

Module M
    Sub Main()
        Dim root As String = Path.Combine(Path.GetTempPath(), "vb_dir_matrix_" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(root)
        Console.WriteLine(Directory.Exists(root))
        Directory.Delete(root)
        Console.WriteLine(Directory.Exists(root))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn directory_info_full_path_and_name() {
    let out = run_vb(
        r#"
Imports System.IO

Module M
    Sub Main()
        Dim root As New DirectoryInfo(Path.Combine(Path.GetTempPath(), "vb_info_" & Guid.NewGuid().ToString("N")))
        root.Create()
        Dim info As New DirectoryInfo(root.FullName)
        Console.WriteLine(info.Exists)
        Console.WriteLine(info.Name.StartsWith("vb_info_"))
        root.Delete()
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn directory_move_renames_directory() {
    let out = run_vb(
        r#"
Imports System.IO

Module M
    Sub Main()
        Dim source As String = Path.Combine(Path.GetTempPath(), "vb_move_src_" & Guid.NewGuid().ToString("N"))
        Dim destination As String = Path.Combine(Path.GetTempPath(), "vb_move_dst_" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(source)
        Directory.Move(source, destination)
        Console.WriteLine(Directory.Exists(source))
        Console.WriteLine(Directory.Exists(destination))
        Directory.Delete(destination)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["False", "True"]);
}

#[test]
fn directory_create_nested_and_list_files() {
    let out = run_vb(
        r#"
Imports System.IO

Module M
    Sub Main()
        Dim root As String = Path.Combine(Path.GetTempPath(), "vb_nested_" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(Path.Combine(root, "a", "b"))
        File.WriteAllText(Path.Combine(root, "a", "b", "x.txt"), "v")
        Dim files As String() = Directory.GetFiles(root, "*.*", SearchOption.AllDirectories)
        Dim dirs As String() = Directory.GetDirectories(root, "*", SearchOption.AllDirectories)
        Console.WriteLine(files.Length)
        Console.WriteLine(dirs.Length)
        Directory.Delete(root, True)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn directory_get_parent_reference() {
    let out = run_vb(
        r#"
Imports System.IO

Module M
    Sub Main()
        Dim root As String = Path.Combine(Path.GetTempPath(), "vb_parent_" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(Path.Combine(root, "child"))
        Dim child As String = Path.Combine(root, "child")
        Dim parent As String = Directory.GetParent(child).FullName
        Console.WriteLine(parent = root)
        Directory.Delete(root, True)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}

#[test]
fn directory_enumerate_files_is_ordered_after_create() {
    let out = run_vb(
        r#"
Imports System.IO

Module M
    Sub Main()
        Dim root As String = Path.Combine(Path.GetTempPath(), "vb_enumerate_" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(root)
        File.WriteAllText(Path.Combine(root, "b.txt"), "2")
        File.WriteAllText(Path.Combine(root, "a.txt"), "1")
        Dim items As String() = Directory.GetFiles(root, "*.txt")
        Console.WriteLine(items.Length)
        Console.WriteLine(Path.GetFileName(items(0)))
        Directory.Delete(root, True)
    End Module
End Module
"#,
    );

    assert_eq!(out, vec!["2", "b.txt"]);
}

#[test]
fn directory_enumerate_directories_returns_empty_for_leaf() {
    let out = run_vb(
        r#"
Imports System.IO

Module M
    Sub Main()
        Dim root As String = Path.Combine(Path.GetTempPath(), "vb_leaf_" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(root)
        Dim subdirs As String() = Directory.GetDirectories(root)
        Console.WriteLine(subdirs.Length)
        Directory.Delete(root, True)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["0"]);
}
