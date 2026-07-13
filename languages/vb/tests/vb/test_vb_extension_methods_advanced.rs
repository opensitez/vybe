use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Extension Methods Advanced
// ═══════════════════════════════════════════════════════════

#[test]
fn extension_methods_generics() {
    let out = run_vb(
        r#"
Imports System.Runtime.CompilerServices
Imports System.Collections.Generic

Module ExtensionMethods
    <Extension()>
    Public Sub PrintItems(Of T)(collection As IEnumerable(Of T))
        For Each item In collection
            Console.WriteLine(item)
        Next
    End Sub
End Module

Module M
    Sub Main()
        Dim list As New List(Of Integer) From { 1, 2, 3 }
        list.PrintItems()
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["1", "2", "3"]);
}
