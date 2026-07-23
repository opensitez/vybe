use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Multidimensional Default Properties (Indexers)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_indexer_2d_matrix_access() {
    let src = r#"
Class SparseMatrix
    Private _data As New System.Collections.Generic.Dictionary(Of String, Double)()

    Default Public Property Item(row As Integer, col As Integer) As Double
        Get
            Dim key As String = row & "," & col
            If _data.ContainsKey(key) Then Return _data(key)
            Return 0.0
        End Get
        Set(value As Double)
            Dim key As String = row & "," & col
            _data(key) = value
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim mat As New SparseMatrix()
        mat(1, 2) = 3.14
        Console.WriteLine(mat(1, 2))
        Console.WriteLine(mat(0, 0))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3.14", "0"]);
}

#[test]
fn test_vb_indexer_overloaded_string_and_int() {
    let src = r#"
Class DataStore
    Private _byInt As New System.Collections.Generic.Dictionary(Of Integer, String)()
    Private _byStr As New System.Collections.Generic.Dictionary(Of String, String)()

    Default Public Property Item(id As Integer) As String
        Get
            Return _byInt(id)
        End Get
        Set(value As String)
            _byInt(id) = value
        End Set
    End Property

    Default Public Property Item(key As String) As String
        Get
            Return _byStr(key)
        End Get
        Set(value As String)
            _byStr(key) = value
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim ds As New DataStore()
        ds(1) = "NumOne"
        ds("A") = "StrA"
        Console.WriteLine(ds(1))
        Console.WriteLine(ds("A"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["NumOne", "StrA"]);
}

#[test]
fn test_vb_indexer_read_only_default_property() {
    let src = r#"
Class ReadOnlyGrid
    Default Public ReadOnly Property Item(row As Integer, col As Integer) As String
        Get
            Return "Cell(" & row & "," & col & ")"
        End Get
    End Property
End Class

Module Program
    Sub Main()
        Dim g As New ReadOnlyGrid()
        Console.WriteLine(g(2, 3))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Cell(2,3)"]);
}
