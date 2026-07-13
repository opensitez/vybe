use super::helpers::run_vb;

#[test]
fn property_default_args() {
    let out = run_vb(
        r#"
Class Cache
    Private _items As New System.Collections.Generic.Dictionary(Of String, String)
    
    ' Default Property allows the object to be indexed directly like an array
    Default Public Property Item(key As String) As String
        Get
            If _items.ContainsKey(key) Then Return _items(key)
            Return Nothing
        End Get
        Set(value As String)
            _items(key) = value
        End Set
    End Property
End Class

Module M
    Sub Main()
        Dim c As New Cache()
        ' Using the default property
        c("A") = "Apple"
        c("B") = "Banana"
        
        Console.WriteLine(c("A"))
        Console.WriteLine(c("B"))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Apple", "Banana"]);
}
