use super::helpers::run_vb;

#[test]
fn default_properties_overloaded() {
    let out = run_vb(
        r#"
Class Matrix
    Private data(2, 2) As Integer
    
    ' Overloaded Default Properties
    Default Public Property Item(row As Integer, col As Integer) As Integer
        Get
            Return data(row, col)
        End Get
        Set(value As Integer)
            data(row, col) = value
        End Set
    End Property
    
    Default Public Property Item(index As String) As Integer
        Get
            Dim parts = index.Split(","c)
            Return data(Integer.Parse(parts(0)), Integer.Parse(parts(1)))
        End Get
        Set(value As Integer)
            Dim parts = index.Split(","c)
            data(Integer.Parse(parts(0)), Integer.Parse(parts(1))) = value
        End Set
    End Property
End Class

Module M
    Sub Main()
        Dim m As New Matrix()
        m(0, 1) = 5
        m("1,2") = 10
        
        Console.WriteLine(m(0, 1))
        Console.WriteLine(m("1,2"))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["5", "10"]);
}
