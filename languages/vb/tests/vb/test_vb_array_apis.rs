use super::helpers::run_vb;

#[test]
fn array_index_of_finds_first_matching_element() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim values As Integer() = {4, 7, 9}
        Console.WriteLine(Array.IndexOf(values, 7))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["1"]);
}

#[test]
fn array_last_index_of_finds_last_match() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim values As Integer() = {1, 2, 1, 3}
        Console.WriteLine(Array.LastIndexOf(values, 1))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["2"]);
}

#[test]
fn array_reverse_reorders_values() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim values As Integer() = {1, 2, 3}
        Array.Reverse(values)
        For Each value As Integer In values
            Console.WriteLine(value)
        Next
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["3", "2", "1"]);
}

#[test]
fn array_clear_resets_range_to_default() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim values As Integer() = {1, 2, 3}
        Array.Clear(values, 1, 2)
        For Each value As Integer In values
            Console.WriteLine(value)
        Next
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["1", "0", "0"]);
}

#[test]
fn array_copy_moves_values_between_arrays() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim source As Integer() = {5, 6, 7}
        Dim target As Integer() = New Integer(2) {}
        Array.Copy(source, target, 3)
        For Each value As Integer In target
            Console.WriteLine(value)
        Next
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["5", "6", "7"]);
}

#[test]
fn array_resize_grows_array_and_preserves_existing_values() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim values As Integer() = {2, 4}
        Array.Resize(values, 4)
        For Each value As Integer In values
            Console.WriteLine(value)
        Next
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["2", "4", "0", "0"]);
}

#[test]
fn array_sort_orders_values_ascending() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim values As Integer() = {4, 1, 3}
        Array.Sort(values)
        For Each value As Integer In values
            Console.WriteLine(value)
        Next
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["1", "3", "4"]);
}

#[test]
fn array_binary_search_returns_position() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim values As Integer() = {1, 3, 5, 7}
        Console.WriteLine(Array.BinarySearch(values, 5))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["2"]);
}

#[test]
fn array_exists_reports_predicate_match() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim values As Integer() = {1, 3, 5}
        Console.WriteLine(Array.Exists(values, Function(value As Integer) value = 3))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}

#[test]
fn array_find_returns_first_matching_value() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim values As Integer() = {2, 4, 5, 8}
        Console.WriteLine(Array.Find(values, Function(value As Integer) value Mod 2 = 1))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["5"]);
}

#[test]
fn array_find_index_returns_position() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim values As Integer() = {2, 4, 5, 8}
        Console.WriteLine(Array.FindIndex(values, Function(value As Integer) value Mod 2 = 1))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["2"]);
}

#[test]
fn array_convert_all_maps_to_strings() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim values As Integer() = {1, 2, 3}
        Dim text As String() = Array.ConvertAll(values, Function(value As Integer) "n" & value)
        For Each part As String In text
            Console.WriteLine(part)
        Next
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["n1", "n2", "n3"]);
}

#[test]
fn array_true_for_all_checks_sequence() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim values As Integer() = {2, 4, 6}
        Console.WriteLine(Array.TrueForAll(values, Function(value As Integer) value Mod 2 = 0))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}

#[test]
fn array_empty_returns_zero_length_array() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine(Array.Empty(Of String)().Length)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["0"]);
}

#[test]
fn array_create_instance_reports_length() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim arr As Array = Array.CreateInstance(GetType(Integer), 3)
        Console.WriteLine(arr.Length)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["3"]);
}

#[test]
fn array_rank_reports_dimensions() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim values As Integer() = {1, 2, 3}
        Console.WriteLine(values.Rank)
        Console.WriteLine(values.GetLength(0))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["1", "3"]);
}

#[test]
fn multidimensional_array_length_by_axis() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim matrix(1, 2) As Integer
        Console.WriteLine(matrix.GetLength(0))
        Console.WriteLine(matrix.GetLength(1))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["2", "3"]);
}

#[test]
fn array_copyto_moves_values_into_target() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim source As Integer() = {9, 8}
        Dim target As Integer() = New Integer(1) {}
        source.CopyTo(target, 0)
        For Each value As Integer In target
            Console.WriteLine(value)
        Next
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["9", "8"]);
}
