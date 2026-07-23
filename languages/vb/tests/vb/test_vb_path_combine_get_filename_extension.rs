use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: System.IO.Path Utilities & Cross-Platform Normalization
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_path_combine_two_paths() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Dim combined = Path.Combine("folder", "file.txt")
        Console.WriteLine(combined.EndsWith("file.txt") AndAlso combined.Contains("folder"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_path_combine_four_paths_overload() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Dim combined = Path.Combine("usr", "local", "bin", "app")
        Console.WriteLine(combined.EndsWith("app"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_path_get_file_name_with_extension() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Dim filename = Path.GetFileName("/var/log/app.log")
        Console.WriteLine(filename)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["app.log"]);
}

#[test]
fn test_vb_path_get_file_name_without_extension() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Dim nameOnly = Path.GetFileNameWithoutExtension("/path/to/archive.tar.gz")
        Console.WriteLine(nameOnly)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["archive.tar"]);
}

#[test]
fn test_vb_path_get_extension() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Dim ext = Path.GetExtension("document.pdf")
        Console.WriteLine(ext)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec![".pdf"]);
}

#[test]
fn test_vb_path_change_extension() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Dim changed = Path.ChangeExtension("image.png", ".jpg")
        Console.WriteLine(changed)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["image.jpg"]);
}

#[test]
fn test_vb_path_change_extension_remove_extension() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Dim noExt = Path.ChangeExtension("script.vb", Nothing)
        Console.WriteLine(noExt)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["script"]);
}

#[test]
fn test_vb_path_get_directory_name() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Dim dirName = Path.GetDirectoryName("/home/user/data.csv")
        Console.WriteLine(dirName.EndsWith("user"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_path_get_path_root() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Dim root = Path.GetPathRoot("/usr/bin")
        Console.WriteLine(root)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["/"]);
}

#[test]
fn test_vb_path_has_extension() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Console.WriteLine(Path.HasExtension("file.txt") & "|" & Path.HasExtension("folder/file"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False"]);
}

#[test]
fn test_vb_path_get_random_file_name() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Dim rndName = Path.GetRandomFileName()
        Console.WriteLine(rndName.Length > 0 AndAlso Path.HasExtension(rndName))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_path_get_temp_file_name() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Dim tempFile = Path.GetTempFileName()
        Console.WriteLine(File.Exists(tempFile))
        File.Delete(tempFile)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_path_get_temp_path() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Dim tempDir = Path.GetTempPath()
        Console.WriteLine(tempDir.Length > 0)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_path_is_path_rooted() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Console.WriteLine(Path.IsPathRooted("/abs/path") & "|" & Path.IsPathRooted("rel/path"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False"]);
}

#[test]
fn test_vb_path_get_relative_path() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Dim rel = Path.GetRelativePath("/home/user", "/home/user/docs/file.txt")
        Console.WriteLine(rel.Replace("\", "/"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["docs/file.txt"]);
}

#[test]
fn test_vb_path_join_array() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Dim joined = Path.Join("a", "b", "c")
        Console.WriteLine(joined.Contains("b"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_path_directory_separator_char() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Dim sep = Path.DirectorySeparatorChar
        Console.WriteLine(sep = "/"c OrElse sep = "\"c)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_path_get_invalid_filename_chars() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Dim invalidChars = Path.GetInvalidFileNameChars()
        Console.WriteLine(invalidChars.Length > 0)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_path_get_invalid_path_chars() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Dim invalidChars = Path.GetInvalidPathChars()
        Console.WriteLine(invalidChars.Length > 0)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_path_trim_ending_directory_separator() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Dim trimmed = Path.TrimEndingDirectorySeparator("/folder/sub/")
        Console.WriteLine(trimmed.EndsWith("sub"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}
