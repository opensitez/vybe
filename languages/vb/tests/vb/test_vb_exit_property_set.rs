use super::helpers::run_vb;

#[test]
fn exit_property_set() {
    let out = run_vb(
        r#"
Class Cache
    Private _val As Integer
    
    Public Property Value As Integer
        Get
            Return _val
        End Get
        Set(val As Integer)
            If val < 0 Then Exit Property
            _val = val
        End Set
    End Property
End Class

Module M
    Sub Main()
        Dim c As New Cache()
        c.Value = -10
        Console.WriteLine(c.Value) ' Should be 0
        c.Value = 20
        Console.WriteLine(c.Value) ' Should be 20
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["0", "20"]);
}
