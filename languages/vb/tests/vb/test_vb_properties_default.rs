use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Default Properties
// ═══════════════════════════════════════════════════════════

#[test]
fn default_property_basic() {
    let out = run_vb(
        r#"
Class StringCollection
    Private _items(10) As String
    
    Default Public Property Item(index As Integer) As String
        Get
            Return _items(index)
        End Get
        Set(value As String)
            _items(index) = value
        End Set
    End Property
End Class

Module M
    Sub Main()
        Dim col As New StringCollection()
        ' Accessing via default property syntax
        col(0) = "First"
        col(1) = "Second"
        
        Console.WriteLine(col(0))
        Console.WriteLine(col(1))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["First", "Second"]);
}

#[test]
fn default_property_multiple_args() {
    let out = run_vb(
        r#"
Class Map2D
    Private _map(5, 5) As Integer
    
    Default Public Property Cell(x As Integer, y As Integer) As Integer
        Get
            Return _map(x, y)
        End Get
        Set(value As Integer)
            _map(x, y) = value
        End Set
    End Property
End Class

Module M
    Sub Main()
        Dim map As New Map2D()
        map(2, 3) = 99
        Console.WriteLine(map(2, 3))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["99"]);
}

#[test]
fn default_property_overloads() {
    let out = run_vb(
        r#"
Class DictionaryMock
    Private _keys(10) As String
    Private _values(10) As String
    Private _count As Integer = 0
    
    Public Sub Add(key As String, value As String)
        _keys(_count) = key
        _values(_count) = value
        _count += 1
    End Sub

    Default Public ReadOnly Property Item(key As String) As String
        Get
            For i As Integer = 0 To _count - 1
                If _keys(i) = key Then Return _values(i)
            Next
            Return "Not Found"
        End Get
    End Property
    
    Default Public ReadOnly Property Item(index As Integer) As String
        Get
            If index >= 0 AndAlso index < _count Then
                Return _values(index)
            End If
            Return "Out of Bounds"
        End Get
    End Property
End Class

Module M
    Sub Main()
        Dim dict As New DictionaryMock()
        dict.Add("Apples", "Red")
        dict.Add("Bananas", "Yellow")
        
        ' Call by string key
        Console.WriteLine(dict("Bananas"))
        
        ' Call by integer index
        Console.WriteLine(dict(0))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Yellow", "Red"]);
}
