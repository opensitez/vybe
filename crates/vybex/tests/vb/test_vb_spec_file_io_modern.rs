use super::helpers::run_vb;
use std::path::{Path, PathBuf};
use uuid::Uuid;

fn vb_path(path: &Path) -> String {
    path.display().to_string().replace('\\', "\\\\")
}

fn temp_root() -> PathBuf {
    std::env::temp_dir().join(format!("vybex_vb_file_io_modern_{}", Uuid::new_v4()))
}

#[test]
fn file_io_modern_spec_writealltext_and_readalltext_roundtrip() {
    let root = temp_root();
    std::fs::create_dir_all(&root).unwrap();
    let file = root.join("note.txt");

    let src = format!(
        r#"
Imports System.IO
Module Program
    Sub Main()
        File.WriteAllText("{file}", "alpha")
        Console.WriteLine(File.ReadAllText("{file}"))
    End Sub
End Module
"#,
        file = vb_path(&file)
    );

    let output = run_vb(&src);
    assert_eq!(output, vec!["alpha"]);
    let _ = std::fs::remove_dir_all(&root);
}

#[test]
fn file_io_modern_spec_appendalltext_and_readalllines() {
    let root = temp_root();
    std::fs::create_dir_all(&root).unwrap();
    let file = root.join("append.txt");

    let src = format!(
        r#"
Imports System.IO
Module Program
    Sub Main()
        File.WriteAllText("{file}", "first")
        File.AppendAllText("{file}", Chr(10) & "second")
        Dim lines() As String = File.ReadAllLines("{file}")
        Console.WriteLine(lines(0))
        Console.WriteLine(lines(1))
    End Sub
End Module
"#,
        file = vb_path(&file)
    );

    let output = run_vb(&src);
    assert_eq!(output, vec!["first", "second"]);
    let _ = std::fs::remove_dir_all(&root);
}

#[test]
fn file_io_modern_spec_exists_copy_move_delete() {
    let root = temp_root();
    std::fs::create_dir_all(&root).unwrap();
    let source = root.join("source.txt");
    let copy = root.join("copy.txt");
    let moved = root.join("moved.txt");

    let src = format!(
        r#"
Imports System.IO
Module Program
    Sub Main()
        File.WriteAllText("{source}", "payload")
        Console.WriteLine(File.Exists("{source}"))
        File.Copy("{source}", "{copy}")
        Console.WriteLine(File.ReadAllText("{copy}"))
        File.Move("{copy}", "{moved}")
        Console.WriteLine(File.Exists("{copy}"))
        Console.WriteLine(File.Exists("{moved}"))
        File.Delete("{moved}")
        Console.WriteLine(File.Exists("{moved}"))
    End Sub
End Module
"#,
        source = vb_path(&source),
        copy = vb_path(&copy),
        moved = vb_path(&moved)
    );

    let output = run_vb(&src);
    assert_eq!(output, vec!["True", "payload", "False", "True", "False"]);
    let _ = std::fs::remove_dir_all(&root);
}

#[test]
fn file_io_modern_spec_path_helpers() {
    let root = temp_root();
    std::fs::create_dir_all(&root).unwrap();
    let nested = root.join("folder").join("sample.txt");

    let src = format!(
        r#"
Imports System.IO
Module Program
    Sub Main()
        Console.WriteLine(Path.GetFileName("{file}"))
        Console.WriteLine(Path.GetExtension("{file}"))
        Console.WriteLine(Path.GetDirectoryName("{file}").EndsWith("folder"))
        Console.WriteLine(Path.ChangeExtension("{file}", ".log").EndsWith("sample.log"))
        Console.WriteLine(Len(Path.GetTempPath()) > 0)
    End Sub
End Module
"#,
        file = vb_path(&nested)
    );

    let output = run_vb(&src);
    assert_eq!(output, vec!["sample.txt", ".txt", "True", "True", "True"]);
    let _ = std::fs::remove_dir_all(&root);
}

#[test]
fn file_io_modern_spec_directory_helpers() {
    let root = temp_root();
    let child_dir = root.join("child");
    let child_file = root.join("item.txt");
    std::fs::create_dir_all(&child_dir).unwrap();
    std::fs::write(&child_file, "payload").unwrap();

    let src = format!(
        r#"
Imports System.IO
Module Program
    Sub Main()
        Console.WriteLine(Directory.Exists("{root}"))
        Dim files() As String = Directory.GetFiles("{root}")
        Dim dirs() As String = Directory.GetDirectories("{root}")
        Console.WriteLine(files.Length)
        Console.WriteLine(dirs.Length)
        Console.WriteLine(Directory.GetCurrentDirectory().Length > 0)
        Directory.Delete("{child}")
        Console.WriteLine(Directory.Exists("{child}"))
    End Sub
End Module
"#,
        root = vb_path(&root),
        child = vb_path(&child_dir)
    );

    let output = run_vb(&src);
    assert_eq!(output, vec!["True", "1", "1", "True", "False"]);
    let _ = std::fs::remove_dir_all(&root);
}

#[test]
fn file_io_modern_spec_streamreader_and_streamwriter() {
    let root = temp_root();
    std::fs::create_dir_all(&root).unwrap();
    let file = root.join("stream.txt");

    let src = format!(
        r#"
Imports System.IO
Module Program
    Sub Main()
        Dim writer As New StreamWriter("{file}")
        writer.WriteLine("line-one")
        writer.WriteLine("line-two")
        writer.Close()

        Dim reader As New StreamReader("{file}")
        Console.WriteLine(reader.ReadLine())
        Console.WriteLine(reader.ReadToEnd().Contains("line-two"))
        Console.WriteLine(reader.EndOfStream)
        reader.Close()
    End Sub
End Module
"#,
        file = vb_path(&file)
    );

    let output = run_vb(&src);
    assert_eq!(output, vec!["line-one", "True", "True"]);
    let _ = std::fs::remove_dir_all(&root);
}
