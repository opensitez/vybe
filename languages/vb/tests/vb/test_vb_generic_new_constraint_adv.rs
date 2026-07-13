use super::helpers::run_vb;

#[test]
fn generic_new_constraint_adv() {
    let out = run_vb(
        r#"
Class Factory(Of T As New)
    Public Function Create() As T
        Return New T()
    End Function
End Class

Class Item
    Public Sub New()
        Console.WriteLine("ItemCreated")
    End Sub
End Class

Module M
    Sub Main()
        Dim f As New Factory(Of Item)()
        f.Create()
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["ItemCreated"]);
}
