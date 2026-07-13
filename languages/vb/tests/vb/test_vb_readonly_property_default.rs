use super::helpers::run_vb;

#[test]
fn readonly_property_default() {
    let out = run_vb(
        r#"
Class Configuration
    ' ReadOnly auto-property with default initializer
    Public ReadOnly Property Version As String = "1.0.0"
    
    ' Property with backing field
    Private _maxRetries As Integer = 3
    Public ReadOnly Property MaxRetries As Integer
        Get
            Return _maxRetries
        End Get
    End Property
End Class

Module M
    Sub Main()
        Dim c As New Configuration()
        Console.WriteLine(c.Version)
        Console.WriteLine(c.MaxRetries)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["1.0.0", "3"]);
}
