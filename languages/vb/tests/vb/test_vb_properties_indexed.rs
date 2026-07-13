use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Indexed Properties
// ═══════════════════════════════════════════════════════════

#[test]
fn indexed_property_basic() {
    let out = run_vb(
        r#"
Class StringGrid
    Private _grid(10) As String
    
    Public Property Cell(index As Integer) As String
        Get
            Return _grid(index)
        End Get
        Set(value As String)
            _grid(index) = value
        End Set
    End Property
End Class

Module M
    Sub Main()
        Dim grid As New StringGrid()
        grid.Cell(5) = "Middle"
        grid.Cell(0) = "Start"
        
        Console.WriteLine(grid.Cell(0))
        Console.WriteLine(grid.Cell(5))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Start", "Middle"]);
}

#[test]
fn indexed_property_multiple_parameters() {
    let out = run_vb(
        r#"
Class Matrix2D
    Private _data(2, 2) As Integer
    
    Public Property Item(x As Integer, y As Integer) As Integer
        Get
            Return _data(x, y)
        End Get
        Set(value As Integer)
            _data(x, y) = value
        End Set
    End Property
End Class

Module M
    Sub Main()
        Dim mat As New Matrix2D()
        mat.Item(1, 2) = 42
        Console.WriteLine(mat.Item(1, 2))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn indexed_property_readonly() {
    let out = run_vb(
        r#"
Class MathTable
    Public ReadOnly Property Multiplier(factor As Integer) As Integer
        Get
            Return factor * 10
        End Get
    End Property
End Class

Module M
    Sub Main()
        Dim table As New MathTable()
        Console.WriteLine(table.Multiplier(5))
        Console.WriteLine(table.Multiplier(9))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["50", "90"]);
}
