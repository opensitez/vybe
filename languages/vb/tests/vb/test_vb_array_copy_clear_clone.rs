use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Array Copy, Clear, and Clone Operations
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_array_copy_1d_range() {
    let src = r#"
Module Program
    Sub Main()
        Dim srcArr As Integer() = {10, 20, 30, 40, 50}
        Dim dstArr(4) As Integer
        Array.Copy(srcArr, 1, dstArr, 2, 3)
        Console.WriteLine(String.Join(",", dstArr))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0,0,20,30,40"]);
}

#[test]
fn test_vb_array_clear_range_primitives() {
    let src = r#"
Module Program
    Sub Main()
        Dim arr As Integer() = {1, 2, 3, 4, 5}
        Array.Clear(arr, 1, 3)
        Console.WriteLine(String.Join(",", arr))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1,0,0,0,5"]);
}

#[test]
fn test_vb_array_clear_range_reference_types() {
    let src = r#"
Module Program
    Sub Main()
        Dim arr As String() = {"a", "b", "c", "d"}
        Array.Clear(arr, 1, 2)
        Console.WriteLine(arr(0))
        Console.WriteLine(arr(1) Is Nothing)
        Console.WriteLine(arr(3))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["a", "True", "d"]);
}

#[test]
fn test_vb_array_constrained_copy_success() {
    let src = r#"
Module Program
    Sub Main()
        Dim srcArr As Integer() = {1, 2, 3}
        Dim dstArr(2) As Integer
        Array.ConstrainedCopy(srcArr, 0, dstArr, 0, 3)
        Console.WriteLine(String.Join(",", dstArr))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1,2,3"]);
}

#[test]
fn test_vb_array_clone_shallow_copy() {
    let src = r#"
Module Program
    Sub Main()
        Dim orig As String() = {"X", "Y", "Z"}
        Dim cloneArr As String() = CType(orig.Clone(), String())
        cloneArr(0) = "W"
        Console.WriteLine(orig(0))
        Console.WriteLine(cloneArr(0))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["X", "W"]);
}

#[test]
fn test_vb_array_empty_singleton() {
    let src = r#"
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
fn test_vb_array_fill_value() {
    let src = r#"
Module Program
    Sub Main()
        Dim arr(4) As Integer
        Array.Fill(arr, 42)
        Console.WriteLine(String.Join(",", arr))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["42,42,42,42,42"]);
}

#[test]
fn test_vb_array_fill_range() {
    let src = r#"
Module Program
    Sub Main()
        Dim arr(4) As String
        Array.Fill(arr, "X", 1, 3)
        Console.WriteLine(arr(0) Is Nothing)
        Console.WriteLine(arr(1))
        Console.WriteLine(arr(3))
        Console.WriteLine(arr(4) Is Nothing)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "X", "X", "True"]);
}

#[test]
fn test_vb_array_resize_expand() {
    let src = r#"
Module Program
    Sub Main()
        Dim arr As Integer() = {10, 20}
        Array.Resize(arr, 4)
        Console.WriteLine(arr.Length)
        Console.WriteLine(String.Join(",", arr))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["4", "10,20,0,0"]);
}

#[test]
fn test_vb_array_resize_shrink() {
    let src = r#"
Module Program
    Sub Main()
        Dim arr As Integer() = {10, 20, 30, 40}
        Array.Resize(arr, 2)
        Console.WriteLine(arr.Length)
        Console.WriteLine(String.Join(",", arr))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2", "10,20"]);
}

#[test]
fn test_vb_array_reverse_1d() {
    let src = r#"
Module Program
    Sub Main()
        Dim arr As Integer() = {1, 2, 3, 4, 5}
        Array.Reverse(arr)
        Console.WriteLine(String.Join(",", arr))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["5,4,3,2,1"]);
}

#[test]
fn test_vb_array_reverse_range() {
    let src = r#"
Module Program
    Sub Main()
        Dim arr As Integer() = {10, 20, 30, 40, 50}
        Array.Reverse(arr, 1, 3)
        Console.WriteLine(String.Join(",", arr))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10,40,30,20,50"]);
}

#[test]
fn test_vb_array_true_for_all_predicate() {
    let src = r#"
Module Program
    Sub Main()
        Dim arr As Integer() = {2, 4, 6, 8}
        Dim isAllEven As Boolean = Array.TrueForAll(arr, Function(x) x Mod 2 = 0)
        Console.WriteLine(isAllEven)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_array_find_and_find_last() {
    let src = r#"
Module Program
    Sub Main()
        Dim arr As Integer() = {1, 5, 8, 12, 15}
        Dim firstEven As Integer = Array.Find(arr, Function(x) x Mod 2 = 0)
        Dim lastEven As Integer = Array.FindLast(arr, Function(x) x Mod 2 = 0)
        Console.WriteLine(firstEven)
        Console.WriteLine(lastEven)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["8", "12"]);
}

#[test]
fn test_vb_array_find_all_matches() {
    let src = r#"
Module Program
    Sub Main()
        Dim arr As Integer() = {1, 2, 3, 4, 5, 6}
        Dim evens As Integer() = Array.FindAll(arr, Function(x) x Mod 2 = 0)
        Console.WriteLine(String.Join(",", evens))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2,4,6"]);
}
