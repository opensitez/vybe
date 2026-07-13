use super::helpers::run_vb;

#[test]
fn property_parameters_indexed() {
    let out = run_vb(
        r#"
Class Cache
    Private data(10) As String
    
    ' Property with parameters (Indexed Property)
    Public Property ItemAt(index As Integer) As String
        Get
            Return data(index)
        End Get
        Set(value As String)
            data(index) = value
        End Set
    End Property
End Class

Module M
    Sub Main()
        Dim c As New Cache()
        c.ItemAt(5) = "Stored"
        Console.WriteLine(c.ItemAt(5))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Stored"]);
}
