use super::helpers::run_vb;

#[test]
fn synclock_statement() {
    let out = run_vb(
        r#"
Module M
    Private _syncObj As New Object()
    
    Sub Main()
        ' SyncLock provides exclusive access to a block of code based on an object lock
        SyncLock _syncObj
            Console.WriteLine("Inside Lock")
        End SyncLock
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Inside Lock"]);
}
