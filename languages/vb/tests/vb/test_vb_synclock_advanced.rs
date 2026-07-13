use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: SyncLock Advanced
// ═══════════════════════════════════════════════════════════

#[test]
fn synclock_advanced() {
    let out = run_vb(
        r#"
Imports System.Threading

Class Counter
    Private _count As Integer = 0
    Private _lockObj As New Object()
    
    Public Sub Increment()
        SyncLock _lockObj
            _count += 1
        End SyncLock
    End Sub
    
    Public ReadOnly Property Count As Integer
        Get
            SyncLock _lockObj
                Return _count
            End SyncLock
        End Get
    End Property
End Class

Module M
    Sub Main()
        Dim c As New Counter()
        
        Dim t1 As New Thread(Sub() c.Increment())
        Dim t2 As New Thread(Sub() c.Increment())
        
        t1.Start()
        t2.Start()
        
        t1.Join()
        t2.Join()
        
        Console.WriteLine(c.Count)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["2"]);
}
