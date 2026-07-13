use super::helpers::run_vb;

#[test]
fn byval_properties() {
    let out = run_vb(
        r#"
Class Cache
    Private _val As Integer
    
    ' Set accessors take arguments. Usually implicitly ByVal.
    Public Property Value As Integer
        Get
            Return _val
        End Get
        Set(ByVal val As Integer)
            _val = val
        End Set
    End Property
End Class

Module M
    Sub Main()
        Dim c As New Cache()
        c.Value = 10
        Console.WriteLine(c.Value)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["10"]);
}
