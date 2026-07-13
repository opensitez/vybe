use super::helpers::run_vb;

#[test]
fn default_property_generic() {
    let out = run_vb(
        r#"
Class Cache(Of T)
    Private _dict As New System.Collections.Generic.Dictionary(Of String, T)()
    
    Default Public Property Item(key As String) As T
        Get
            If _dict.ContainsKey(key) Then Return _dict(key)
            Return Nothing
        End Get
        Set(value As T)
            _dict(key) = value
        End Set
    End Property
End Class

Module M
    Sub Main()
        Dim c As New Cache(Of Integer)()
        c("A") = 100
        Console.WriteLine(c("A"))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["100"]);
}
