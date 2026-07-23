use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: WeakReference & Weak Event Pattern
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_weak_reference_alive_and_target() {
    let src = r#"
Imports System

Class TargetData
    Public Property Value As String = "Alive"
End Class

Module Program
    Sub Main()
        Dim obj As New TargetData()
        Dim weak As New WeakReference(obj)
        Console.WriteLine(weak.IsAlive)
        Dim retrieved As TargetData = CType(weak.Target, TargetData)
        Console.WriteLine(retrieved.Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "Alive"]);
}

#[test]
fn test_vb_weak_reference_generic_t() {
    let src = r#"
Imports System

Class DataContainer
    Public Tag As Integer = 42
End Class

Module Program
    Sub Main()
        Dim obj As New DataContainer()
        Dim weak As New WeakReference(Of DataContainer)(obj)
        Dim target As DataContainer = Nothing
        Dim isAlive As Boolean = weak.TryGetTarget(target)
        Console.WriteLine(isAlive)
        Console.WriteLine(target.Tag)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "42"]);
}
