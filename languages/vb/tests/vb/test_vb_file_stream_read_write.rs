use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: File Class Text Operations (WriteAllText, ReadAllText)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_file_write_read_lines() {
    let src = r#"
Imports System.IO

Module Program
    Sub Main()
        Dim tempPath As String = Path.GetTempFileName()
        Try
            File.WriteAllLines(tempPath, New String() {"LineA", "LineB"})
            Dim lines As String() = File.ReadAllLines(tempPath)
            Console.WriteLine(lines.Length)
            Console.WriteLine(lines(0) & "," & lines(1))
        Finally
            If File.Exists(tempPath) Then File.Delete(tempPath)
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2", "LineA,LineB"]);
}
