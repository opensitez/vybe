use super::helpers::run_vb;

#[test]
fn generic_constraint_class() {
    let out = run_vb(
        r#"
' Generic constraint As Class requires T to be a reference type
Class ReferenceCache(Of T As Class)
    Public Property Item As T
End Class

Module M
    Sub Main()
        Dim c As New ReferenceCache(Of String)()
        c.Item = "Hello"
        Console.WriteLine(c.Item)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Hello"]);
}
