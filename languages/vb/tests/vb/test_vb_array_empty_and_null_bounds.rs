use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Array.Empty, Rank, Lower/Upper Bounds & Length Edge Cases
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_array_empty_of_t_singleton() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim empty1 As Integer() = Array.Empty(Of Integer)()
        Dim empty2 As Integer() = Array.Empty(Of Integer)()
        Console.WriteLine(empty1.Length)
        Console.WriteLine(Object.ReferenceEquals(empty1, empty2))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0", "True"]);
}

#[test]
fn test_vb_array_get_lower_bound_1d() {
    let src = r#"
Module Program
    Sub Main()
        Dim arr(5) As Integer
        Console.WriteLine(arr.GetLowerBound(0))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0"]);
}

#[test]
fn test_vb_array_get_upper_bound_1d() {
    let src = r#"
Module Program
    Sub Main()
        Dim arr(5) As Integer
        Console.WriteLine(arr.GetUpperBound(0))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["5"]);
}

#[test]
fn test_vb_array_get_length_dimension() {
    let src = r#"
Module Program
    Sub Main()
        Dim matrix(2, 4) As Double
        Console.WriteLine(matrix.GetLength(0))
        Console.WriteLine(matrix.GetLength(1))
        Console.WriteLine(matrix.Length)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3", "5", "15"]);
}

#[test]
fn test_vb_array_rank_1d_2d_3d() {
    let src = r#"
Module Program
    Sub Main()
        Dim a1(2) As Integer
        Dim a2(2, 2) As Integer
        Dim a3(1, 1, 1) As Integer
        Console.WriteLine(a1.Rank & "," & a2.Rank & "," & a3.Rank)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1,2,3"]);
}

#[test]
fn test_vb_array_get_lower_upper_bounds_multidimensional() {
    let src = r#"
Module Program
    Sub Main()
        Dim grid(3, 7) As String
        Console.WriteLine(grid.GetLowerBound(0) & " To " & grid.GetUpperBound(0))
        Console.WriteLine(grid.GetLowerBound(1) & " To " & grid.GetUpperBound(1))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0 To 3", "0 To 7"]);
}

#[test]
fn test_vb_array_empty_string_array() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim emptyStr As String() = Array.Empty(Of String)()
        Console.WriteLine(emptyStr.Length)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0"]);
}

#[test]
fn test_vb_array_is_fixed_size_is_read_only() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim arr As Integer() = {1, 2, 3}
        Console.WriteLine(arr.IsFixedSize)
        Console.WriteLine(arr.IsReadOnly)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "False"]);
}

#[test]
fn test_vb_array_is_synchronised_sync_root() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim arr As Integer() = {10, 20}
        Console.WriteLine(arr.IsSynchronized)
        Console.WriteLine(arr.SyncRoot IsNot Nothing)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False", "True"]);
}

#[test]
fn test_vb_array_get_value_1d() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim arr As String() = {"Alpha", "Beta", "Gamma"}
        Dim val As Object = arr.GetValue(1)
        Console.WriteLine(val.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Beta"]);
}

#[test]
fn test_vb_array_set_value_1d() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim arr As String() = {"Alpha", "Beta", "Gamma"}
        arr.SetValue("Delta", 1)
        Console.WriteLine(arr(1))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Delta"]);
}

#[test]
fn test_vb_array_get_value_2d() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim grid(,) As Integer = {{1, 2}, {3, 4}}
        Dim val As Object = grid.GetValue(1, 0)
        Console.WriteLine(val)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3"]);
}

#[test]
fn test_vb_array_set_value_2d() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim grid(,) As Integer = {{1, 2}, {3, 4}}
        grid.SetValue(99, 1, 0)
        Console.WriteLine(grid(1, 0))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["99"]);
}

#[test]
fn test_vb_array_create_instance_type_length() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim arr As Array = Array.CreateInstance(GetType(String), 3)
        arr.SetValue("Item0", 0)
        arr.SetValue("Item1", 1)
        Console.WriteLine(arr.Length)
        Console.WriteLine(arr.GetValue(0) & "," & arr.GetValue(1))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3", "Item0,Item1"]);
}

#[test]
fn test_vb_array_create_instance_multidimensional() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim lengths As Integer() = {2, 3}
        Dim grid As Array = Array.CreateInstance(GetType(Integer), lengths)
        grid.SetValue(42, 1, 2)
        Console.WriteLine(grid.Rank)
        Console.WriteLine(grid.GetValue(1, 2))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2", "42"]);
}

#[test]
fn test_vb_array_long_length_property() {
    let src = r#"
Module Program
    Sub Main()
        Dim arr(9) As Byte
        Console.WriteLine(arr.LongLength)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10"]);
}

#[test]
fn test_vb_array_clone_reference_types_shallow() {
    let src = r#"
Class Container
    Public Tag As String
    Public Sub New(t As String)
        Tag = t
    End Sub
End Class

Module Program
    Sub Main()
        Dim orig As Container() = {New Container("A")}
        Dim cloned As Container() = CType(orig.Clone(), Container())
        cloned(0).Tag = "Modified"
        Console.WriteLine(orig(0).Tag)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Modified"]);
}

#[test]
fn test_vb_array_clone_value_types_independent() {
    let src = r#"
Module Program
    Sub Main()
        Dim orig As Integer() = {10, 20}
        Dim cloned As Integer() = CType(orig.Clone(), Integer())
        cloned(0) = 99
        Console.WriteLine(orig(0) & ":" & cloned(0))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10:99"]);
}

#[test]
fn test_vb_array_clear_subset_range() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim arr As Integer() = {10, 20, 30, 40, 50}
        ' Clear 2 items starting from index 1
        Array.Clear(arr, 1, 2)
        Console.WriteLine(String.Join(",", arr))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10,0,0,40,50"]);
}

#[test]
fn test_vb_array_clear_all_items() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim words As String() = {"A", "B", "C"}
        Array.Clear(words, 0, words.Length)
        Console.WriteLine((words(0) Is Nothing) & "," & (words(1) Is Nothing) & "," & (words(2) Is Nothing))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True,True,True"]);
}
