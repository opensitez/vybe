use super::helpers::run_vb;

#[test]
fn array_indexing_matrix_one_based_default_bounds() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim values(1 To 5) As Integer
        For i As Integer = 1 To 5
            values(i) = i * i
        Next

        Console.WriteLine(values.Length)
        Console.WriteLine(values(1))
        Console.WriteLine(values(5))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["5", "1", "25"]);
}

#[test]
fn array_indexing_matrix_zero_based_and_negative_lower_bound() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim values(-2 To 2) As Integer
        values(-2) = -2
        values(-1) = -1
        values(0) = 0
        values(1) = 1
        values(2) = 2

        Console.WriteLine(values.GetLowerBound(0))
        Console.WriteLine(values.GetUpperBound(0))
        Console.WriteLine(values(1) + values(-1))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["-2", "2", "0"]);
}

#[test]
fn array_indexing_matrix_value_assignment_aliases_reference() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim source() As Integer = {1, 2, 3}
        Dim alias() As Integer = source

        alias(1) = 42
        Console.WriteLine(source(1))
        Console.WriteLine(alias.Length)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["42", "3"]);
}

#[test]
fn array_indexing_matrix_redim_preserve_keeps_prefix_values() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim values() As Integer = {1, 2, 3}
        ReDim Preserve values(5)

        values(3) = 4
        values(4) = 5
        values(5) = 6

        Dim sum As Integer = values(0) + values(1) + values(2) + values(3) + values(4) + values(5)
        Console.WriteLine(sum)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["21"]);
}

#[test]
fn array_indexing_matrix_foreach_is_value_ordered() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim values As Integer() = {4, 1, 7, 0}
        Dim ordered As New System.Text.StringBuilder()

        For Each value As Integer In values
            ordered.Append(value).Append(",")
        Next

        Console.WriteLine(ordered.ToString())
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["4,1,7,0,"]);
}

#[test]
fn array_indexing_matrix_clear_and_copy_range() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim values() As Integer = {1, 2, 3, 4, 5}
        Dim left() As Integer = New Integer(2) {}

        Array.Copy(values, 1, left, 0, 3)
        Array.Clear(values, 3, 2)

        Dim c1 As String = String.Join("|", left)
        Dim c2 As String = String.Join("|", values)
        Console.WriteLine(c1)
        Console.WriteLine(c2)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["2|3|4", "1|2|3|0|0"]);
}

#[test]
fn array_indexing_matrix_find_boundaries_with_get_lower_upper_bound() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim values() As Integer = {10, 20, 30}

        Console.WriteLine(values.GetLowerBound(0))
        Console.WriteLine(values.GetUpperBound(0))
        Console.WriteLine(values(values.GetLowerBound(0)))
        Console.WriteLine(values(values.GetUpperBound(0)))
    End Module
"#,
    );

    assert_eq!(out, vec!["0", "2", "10", "30"]);
}
