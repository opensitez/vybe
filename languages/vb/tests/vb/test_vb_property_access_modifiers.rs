use super::helpers::run_vb;

#[test]
fn property_access_modifiers() {
    let out = run_vb(
        r#"
Class Counter
    Private _count As Integer
    
    ' Property is public, but Set is private
    Public Property Count As Integer
        Get
            Return _count
        End Get
        Private Set(value As Integer)
            _count = value
        End Set
    End Property
    
    Public Sub Increment()
        Count += 1
    End Sub
End Class

Module M
    Sub Main()
        Dim c As New Counter()
        c.Increment()
        Console.WriteLine(c.Count)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["1"]);
}
