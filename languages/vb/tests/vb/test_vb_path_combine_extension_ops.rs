use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: System.IO.Path Helpers
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_path_combine_and_extension() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Dim p As String = Path.Combine("folder", "sub", "file.txt")
        Console.WriteLine(Path.GetExtension(p))
        Console.WriteLine(Path.GetFileNameWithoutExtension(p))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec![".txt", "file"]);
}

#[test]
fn test_vb_path_change_extension() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Dim p As String = "data.csv"
        Dim newP As String = Path.ChangeExtension(p, "json")
        Console.WriteLine(newP)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["data.json"]);
}
