use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Default Properties (Advanced)
// ═══════════════════════════════════════════════════════════

#[test]
fn default_properties_multiple_parameters() {
    let out = run_vb(
        r#"
Class Matrix
    Private data(2, 2) As Integer
    
    ' Default property can have multiple parameters
    Default Public Property Item(row As Integer, col As Integer) As Integer
        Get
            Return data(row, col)
        End Get
        Set(value As Integer)
            data(row, col) = value
        End Set
    End Property
End Class

Module M
    Sub Main()
        Dim m As New Matrix()
        
        ' Calling Default Property with multiple arguments
        m(1, 2) = 42
        Console.WriteLine(m(1, 2))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["42"]);
}
