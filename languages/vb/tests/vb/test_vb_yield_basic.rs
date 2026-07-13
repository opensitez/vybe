use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Yield Return (Basic Iterators)
// ═══════════════════════════════════════════════════════════

#[test]
fn yield_return_basic() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Iterator Function GetNumbers() As IEnumerable(Of Integer)
        Yield 1
        Yield 2
        Yield 3
    End Function

    Sub Main()
        For Each n In GetNumbers()
            Console.WriteLine(n)
        Next
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["1", "2", "3"]);
}
