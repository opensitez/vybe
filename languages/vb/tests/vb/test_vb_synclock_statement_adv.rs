use super::helpers::run_vb;

#[test]
fn synclock_statement_adv() {
    let out = run_vb(
        r#"
Module M
    Private lockObj As New Object()

    Sub Main()
        ' SyncLock provides exclusive lock for the block
        SyncLock lockObj
            Console.WriteLine("Locked")
        End SyncLock
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Locked"]);
}
