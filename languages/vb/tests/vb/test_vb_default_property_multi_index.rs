use super::helpers::run_vb;

#[test]
fn default_property_multi_index() {
    let out = run_vb(
        r#"
Class Matrix
    Private data(10, 10) As Integer
    
    Default Public Property Item(x As Integer, y As Integer) As Integer
        Get
            Return data(x, y)
        End Get
        Set(value As Integer)
            data(x, y) = value
        End Set
    End Property
End Class

Module M
    Sub Main()
        Dim m As New Matrix()
        m(5, 5) = 42
        Console.WriteLine(m(5, 5))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["42"]);
}
