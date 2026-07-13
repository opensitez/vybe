use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Modules (Extension Methods)
// ═══════════════════════════════════════════════════════════

#[test]
fn module_extension_methods() {
    let out = run_vb(
        r#"
Imports System.Runtime.CompilerServices

Module StringExtensions
    <Extension()>
    Public Function WordCount(str As String) As Integer
        Return str.Split(New Char() {" "c}, StringSplitOptions.RemoveEmptyEntries).Length
    End Function
End Module

Module M
    Sub Main()
        Dim text As String = "Hello world from VB.NET"
        ' Calling extension method like an instance method
        Console.WriteLine(text.WordCount())
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["4"]);
}
