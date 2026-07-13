use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Exception Filters (When)
// ═══════════════════════════════════════════════════════════

#[test]
fn exception_filters_when() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Try
            Throw New System.InvalidOperationException("Test 1")
        Catch ex As Exception When ex.Message.Contains("2")
            Console.WriteLine("Caught 2")
        Catch ex As Exception When ex.Message.Contains("1")
            Console.WriteLine("Caught 1")
        Catch ex As Exception
            Console.WriteLine("Caught other")
        End Try
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Caught 1"]);
}
