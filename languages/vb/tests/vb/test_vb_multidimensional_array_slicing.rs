use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Multidimensional Array Operations & Bounds
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_array_2d_bounds_lbound_ubound() {
    let src = r#"
Module Program
    Sub Main()
        Dim arr(2, 4) As Integer
        Console.WriteLine(arr.GetLowerBound(0))
        Console.WriteLine(arr.GetUpperBound(0))
        Console.WriteLine(arr.GetLowerBound(1))
        Console.WriteLine(arr.GetUpperBound(1))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0", "2", "0", "4"]);
}

#[test]
fn test_vb_array_2d_rank_length() {
    let src = r#"
Module Program
    Sub Main()
        Dim arr(2, 3) As String
        Console.WriteLine(arr.Rank)
        Console.WriteLine(arr.Length)
        Console.WriteLine(arr.GetLength(0))
        Console.WriteLine(arr.GetLength(1))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2", "12", "3", "4"]);
}

#[test]
fn test_vb_array_3d_element_access() {
    let src = r#"
Module Program
    Sub Main()
        Dim cube(1, 1, 1) As Integer
        cube(0, 0, 0) = 10
        cube(1, 1, 1) = 99
        Console.WriteLine(cube(0, 0, 0))
        Console.WriteLine(cube(1, 1, 1))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10", "99"]);
}

#[test]
fn test_vb_array_2d_nested_loop_row_major() {
    let src = r#"
Module Program
    Sub Main()
        Dim matrix(,) As Integer = {{1, 2}, {3, 4}}
        Dim sum As Integer = 0
        For i As Integer = 0 To matrix.GetUpperBound(0)
            For j As Integer = 0 To matrix.GetUpperBound(1)
                sum += matrix(i, j)
            Next
        Next
        Console.WriteLine(sum)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10"]);
}

#[test]
fn test_vb_array_2d_initializer_syntax() {
    let src = r#"
Module Program
    Sub Main()
        Dim grid As String(,) = {{"a", "b"}, {"c", "d"}}
        Console.WriteLine(grid(0, 1))
        Console.WriteLine(grid(1, 0))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["b", "c"]);
}

#[test]
fn test_vb_array_getvalue_setvalue_dynamic() {
    let src = r#"
Module Program
    Sub Main()
        Dim arr As Array = Array.CreateInstance(GetType(Integer), 2, 3)
        arr.SetValue(42, 1, 2)
        Dim val As Integer = CInt(arr.GetValue(1, 2))
        Console.WriteLine(val)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["42"]);
}

#[test]
fn test_vb_array_2d_clone() {
    let src = r#"
Module Program
    Sub Main()
        Dim orig(,) As Integer = {{10, 20}, {30, 40}}
        Dim copy As Integer(,) = CType(orig.Clone(), Integer(,))
        copy(0, 0) = 999
        Console.WriteLine(orig(0, 0))
        Console.WriteLine(copy(0, 0))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10", "999"]);
}

#[test]
fn test_vb_array_flatten_2d_to_1d() {
    let src = r#"
Module Program
    Sub Main()
        Dim grid(,) As Integer = {{1, 2}, {3, 4}}
        Dim flat(grid.Length - 1) As Integer
        Dim idx As Integer = 0
        For Each val In grid
            flat(idx) = val
            idx += 1
        Next
        Console.WriteLine(String.Join(",", flat))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1,2,3,4"]);
}

#[test]
fn test_vb_array_2d_row_copy() {
    let src = r#"
Module Program
    Sub Main()
        Dim matrix(,) As Integer = {{10, 20, 30}, {40, 50, 60}}
        Dim row1(2) As Integer
        Buffer.BlockCopy(matrix, 0, row1, 0, 3 * sizeof(Integer))
        Console.WriteLine(String.Join(",", row1))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10,20,30"]);
}

#[test]
fn test_vb_array_4d_bounds() {
    let src = r#"
Module Program
    Sub Main()
        Dim arr(1, 2, 3, 4) As Double
        Console.WriteLine(arr.Rank)
        Console.WriteLine(arr.GetLength(3))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["4", "5"]);
}

#[test]
fn test_vb_array_2d_reference_types() {
    let src = r#"
Module Program
    Sub Main()
        Dim names(1, 1) As String
        names(0, 0) = "Alice"
        names(1, 1) = "Bob"
        Console.WriteLine(names(0, 0))
        Console.WriteLine(names(0, 1) Is Nothing)
        Console.WriteLine(names(1, 1))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Alice", "True", "Bob"]);
}

#[test]
fn test_vb_array_2d_fill_diagonal() {
    let src = r#"
Module Program
    Sub Main()
        Dim identity(2, 2) As Integer
        For i As Integer = 0 To 2
            identity(i, i) = 1
        Next
        Console.WriteLine(identity(0, 0))
        Console.WriteLine(identity(0, 1))
        Console.WriteLine(identity(1, 1))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1", "0", "1"]);
}

#[test]
fn test_vb_array_2d_out_of_bounds_exception() {
    let src = r#"
Module Program
    Sub Main()
        Try
            Dim arr(1, 1) As Integer
            Dim x As Integer = arr(2, 0)
            Console.WriteLine(x)
        Catch ex As IndexOutOfRangeException
            Console.WriteLine("IndexOutOfRangeException")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["IndexOutOfRangeException"]);
}

#[test]
fn test_vb_array_2d_transpose() {
    let src = r#"
Module Program
    Sub Main()
        Dim orig(,) As Integer = {{1, 2, 3}, {4, 5, 6}}
        Dim trans(2, 1) As Integer
        For r As Integer = 0 To 1
            For c As Integer = 0 To 2
                trans(c, r) = orig(r, c)
            Next
        Next
        Console.WriteLine(trans(0, 1))
        Console.WriteLine(trans(2, 0))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["4", "3"]);
}

#[test]
fn test_vb_array_2d_is_fixed_size_read_only() {
    let src = r#"
Module Program
    Sub Main()
        Dim arr(1, 1) As Integer
        Console.WriteLine(arr.IsFixedSize)
        Console.WriteLine(arr.IsReadOnly)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "False"]);
}
