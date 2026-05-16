use super::helpers::run_vb;
use std::path::{Path, PathBuf};
use uuid::Uuid;

fn vb_path(path: &Path) -> String {
    path.display().to_string().replace('\\', "\\\\")
}

fn temp_root() -> PathBuf {
    std::env::temp_dir().join(format!("vybex_vb_file_io_legacy_{}", Uuid::new_v4()))
}

#[test]
fn file_io_legacy_spec_open_print_and_close_write_data() {
    let root = temp_root();
    std::fs::create_dir_all(&root).unwrap();
    let file = root.join("print.txt");

    let src = format!(
        r#"
Imports System.IO
Module Program
    Sub Main()
        Dim handle As Integer = FreeFile()
        Open "{file}" For Output As #handle
        Print #handle, "hello from print"
        Close #handle
        Console.WriteLine(File.ReadAllText("{file}").Contains("hello from print"))
    End Sub
End Module
"#,
        file = vb_path(&file)
    );

    let output = run_vb(&src);
    assert_eq!(output, vec!["True"]);
    let _ = std::fs::remove_dir_all(&root);
}

#[test]
fn file_io_legacy_spec_write_and_input_roundtrip() {
    let root = temp_root();
    std::fs::create_dir_all(&root).unwrap();
    let file = root.join("write_input.txt");

    let src = format!(
        r#"
Module Program
    Sub Main()
        Dim handle As Integer = FreeFile()
        Open "{file}" For Output As #handle
        Write #handle, "alpha", 42
        Close #handle

        handle = FreeFile()
        Open "{file}" For Input As #handle
        Dim text As String
        Dim number As Integer
        Input #handle, text, number
        Close #handle

        Console.WriteLine(text)
        Console.WriteLine(number)
    End Sub
End Module
"#,
        file = vb_path(&file)
    );

    let output = run_vb(&src);
    assert_eq!(output, vec!["alpha", "42"]);
    let _ = std::fs::remove_dir_all(&root);
}

#[test]
fn file_io_legacy_spec_dir_kill_and_filecopy() {
    let root = temp_root();
    std::fs::create_dir_all(&root).unwrap();
    let source = root.join("source.txt");
    let copy = root.join("copy.txt");
    std::fs::write(&source, "payload").unwrap();

    let src = format!(
        r#"
Imports System.IO
Module Program
    Sub Main()
        Console.WriteLine(Dir("{source}"))
        FileCopy("{source}", "{copy}")
        Console.WriteLine(File.ReadAllText("{copy}"))
        Kill("{copy}")
        Console.WriteLine(File.Exists("{copy}"))
    End Sub
End Module
"#,
        source = vb_path(&source),
        copy = vb_path(&copy)
    );

    let output = run_vb(&src);
    assert_eq!(output, vec!["source.txt", "payload", "False"]);
    let _ = std::fs::remove_dir_all(&root);
}

#[test]
fn file_io_legacy_spec_mkdir_chdir_curdir_and_rmdir() {
    let root = temp_root();
    std::fs::create_dir_all(&root).unwrap();
    let child = root.join("child");

    let src = format!(
        r#"
Imports System.IO
Module Program
    Sub Main()
        MkDir("{child}")
        Console.WriteLine(Directory.Exists("{child}"))
        ChDir("{child}")
        Console.WriteLine(CurDir().EndsWith("child"))
        ChDir("{root}")
        RmDir("{child}")
        Console.WriteLine(Directory.Exists("{child}"))
    End Sub
End Module
"#,
        child = vb_path(&child),
        root = vb_path(&root)
    );

    let output = run_vb(&src);
    assert_eq!(output, vec!["True", "True", "False"]);
    let _ = std::fs::remove_dir_all(&root);
}

#[test]
fn file_io_legacy_spec_file_metadata_helpers() {
    let root = temp_root();
    std::fs::create_dir_all(&root).unwrap();
    let file = root.join("meta.txt");
    std::fs::write(&file, "abc").unwrap();

    let src = format!(
        r#"
Module Program
    Sub Main()
        Dim attr As Integer = GetAttr("{file}")
        SetAttr("{file}", attr)
        Console.WriteLine(FileLen("{file}"))
        Console.WriteLine(IsDate(CStr(FileDateTime("{file}"))))
        Console.WriteLine(attr >= 0)
    End Sub
End Module
"#,
        file = vb_path(&file)
    );

    let output = run_vb(&src);
    assert_eq!(output, vec!["3", "True", "True"]);
    let _ = std::fs::remove_dir_all(&root);
}

#[test]
fn file_io_legacy_spec_eof_lof_and_loc_report_progress() {
    let root = temp_root();
    std::fs::create_dir_all(&root).unwrap();
    let file = root.join("handle.txt");

    let src = format!(
        r#"
Module Program
    Sub Main()
        Dim handle As Integer = FreeFile()
        Open "{file}" For Output As #handle
        Write #handle, "abc"
        Close #handle

        handle = FreeFile()
        Open "{file}" For Input As #handle
        Console.WriteLine(LOF(handle) > 0)
        Console.WriteLine(LOC(handle) >= 0)
        Console.WriteLine(EOF(handle))
        Dim text As String
        Input #handle, text
        Console.WriteLine(EOF(handle))
        Close #handle
    End Sub
End Module
"#,
        file = vb_path(&file)
    );

    let output = run_vb(&src);
    assert_eq!(output, vec!["True", "True", "False", "True"]);
    let _ = std::fs::remove_dir_all(&root);
}
