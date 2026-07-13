use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: SyncLock Block
// ═══════════════════════════════════════════════════════════

#[test]
fn synclock_basic() {
    let out = run_vb(
        r#"
Class Resource
    Public Value As Integer = 0
End Class

Module M
    Private _lockObj As New Object()
    
    Sub Main()
        Dim res As New Resource()
        
        ' Note: actual multithreading test is complex in basic VM output,
        ' but we can test that the SyncLock syntax parses and executes the block.
        SyncLock _lockObj
            res.Value = res.Value + 10
            Console.WriteLine("Locked and loaded")
        End SyncLock
        
        Console.WriteLine(res.Value)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Locked and loaded", "10"]);
}

#[test]
fn synclock_nested() {
    let out = run_vb(
        r#"
Module M
    Private _lockA As New Object()
    Private _lockB As New Object()
    
    Sub Main()
        SyncLock _lockA
            Console.WriteLine("Lock A acquired")
            SyncLock _lockB
                Console.WriteLine("Lock B acquired")
            End SyncLock
        End SyncLock
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Lock A acquired", "Lock B acquired"]);
}
