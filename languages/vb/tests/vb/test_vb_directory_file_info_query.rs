use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: DirectoryInfo & FileInfo Object Model
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_file_info_properties() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Dim tempPath As String = Path.GetTempFileName()
        Try
            Dim fi As New FileInfo(tempPath)
            Console.WriteLine(fi.Exists)
            Console.WriteLine(fi.Length)
        Finally
            If File.Exists(tempPath) Then File.Delete(tempPath)
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "0"]);
}
