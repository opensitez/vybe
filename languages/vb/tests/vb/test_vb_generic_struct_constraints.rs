use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Generic Constraints on Structure Types
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_generic_struct_with_constraints() {
    let src = r#"
Imports System

Structure Pair(Of TKey As IComparable, TValue)
    Public Key As TKey
    Public Value As TValue

    Public Sub New(k As TKey, v As TValue)
        Me.Key = k
        Me.Value = v
    End Sub
End Structure

Module Program
    Sub Main()
        Dim p As New Pair(Of Integer, String)(10, "Ten")
        Console.WriteLine(p.Key & ":" & p.Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10:Ten"]);
}
