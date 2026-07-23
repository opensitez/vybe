use super::helpers::run_vb;

#[test]
fn array_2d_matrix_construct_and_read() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim grid(1 To 2, 1 To 3) As Integer
        grid(1, 1) = 10
        grid(1, 2) = 20
        grid(1, 3) = 30
        grid(2, 1) = 40
        grid(2, 2) = 50
        grid(2, 3) = 60

        Console.WriteLine(grid.GetLength(0))
        Console.WriteLine(grid.GetLength(1))
        Console.WriteLine(grid(1, 1) + grid(1, 2) + grid(2, 3))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["2", "3", "110"]);
}

#[test]
fn array_2d_matrix_zero_based_bounds_and_indexing() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim grid(0 To 1, 0 To 1) As String
        grid(0, 0) = "a"
        grid(0, 1) = "b"
        grid(1, 0) = "c"
        grid(1, 1) = "d"

        Console.WriteLine(grid.GetLowerBound(0))
        Console.WriteLine(grid.GetLowerBound(1))
        Console.WriteLine(grid(1, 0) & grid(0, 1))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["0", "0", "cb"]);
}

#[test]
fn array_2d_matrix_for_each_like_iteration_not_supported() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim matrix As Integer(,) = New Integer(,) {{1, 2, 3}, {4, 5, 6}}
        Dim first As Integer = matrix(0, 0)
        Dim last As Integer = matrix(1, 2)
        Console.WriteLine(first)
        Console.WriteLine(last)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["1", "6"]);
}

#[test]
fn array_2d_matrix_sum_rows_and_columns() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim m(2, 1) As Integer
        m(0, 0) = 1
        m(0, 1) = 2
        m(1, 0) = 3
        m(1, 1) = 4
        m(2, 0) = 5
        m(2, 1) = 6

        Dim rowSums(2) As Integer
        For r As Integer = m.GetLowerBound(0) To m.GetUpperBound(0)
            For c As Integer = m.GetLowerBound(1) To m.GetUpperBound(1)
                rowSums(r) += m(r, c)
            Next
        Next

        Console.WriteLine(rowSums(0))
        Console.WriteLine(rowSums(1))
        Console.WriteLine(rowSums(2))

        Dim col0 As Integer = m(0, 0) + m(1, 0) + m(2, 0)
        Dim col1 As Integer = m(0, 1) + m(1, 1) + m(2, 1)
        Console.WriteLine(col0)
        Console.WriteLine(col1)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["3", "7", "11", "9", "12"]);
}

#[test]
fn array_2d_matrix_reshape_like_copy_pattern() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim source(1, 2) As Integer
        Dim value As Integer = 1
        For i As Integer = source.GetLowerBound(0) To source.GetUpperBound(0)
            For j As Integer = source.GetLowerBound(1) To source.GetUpperBound(1)
                source(i, j) = value
                value += 1
            Next
        Next

        Dim copied(1, 2) As Integer
        Array.Copy(source, copied, source.Length)

        Console.WriteLine(copied(0, 0))
        Console.WriteLine(copied(0, 2))
        Console.WriteLine(copied(1, 1))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["1", "3", "5"]);
}

#[test]
fn array_2d_matrix_non_square_access_and_bounds() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim matrix(1 To 2, 1 To 4) As Integer
        Console.WriteLine(matrix.Rank)
        Console.WriteLine(matrix.GetUpperBound(0))
        Console.WriteLine(matrix.GetUpperBound(1))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["2", "2", "4"]);
}

#[test]
fn array_2d_matrix_fill_with_nested_loops_then_scan() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim m(0 To 3, 0 To 1) As Integer
        Dim seed As Integer = 1
        For r As Integer = m.GetLowerBound(0) To m.GetUpperBound(0)
            For c As Integer = m.GetLowerBound(1) To m.GetUpperBound(1)
                m(r, c) = seed
                seed += 1
            Next
        Next

        Dim total As Integer = 0
        For r As Integer = m.GetLowerBound(0) To m.GetUpperBound(0)
            For c As Integer = m.GetLowerBound(1) To m.GetUpperBound(1)
                total += m(r, c)
            Next
        Next

        Console.WriteLine(total)
        Console.WriteLine(seed)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["36", "9"]);
}

#[test]
fn array_2d_matrix_jagged_to_rectangular_distinction() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim rect(1, 1) As Integer
        rect(0, 0) = 1
        rect(0, 1) = 2
        rect(1, 0) = 3
        rect(1, 1) = 4

        Dim rows()() As Integer = {New Integer() {1, 2}, New Integer() {3, 4, 5}}
        Dim totalRect As Integer = rect(0, 0) + rect(0, 1) + rect(1, 0) + rect(1, 1)
        Dim totalJagged As Integer = rows(0)(0) + rows(0)(1) + rows(1)(0) + rows(1)(1)

        Console.WriteLine(totalRect)
        Console.WriteLine(totalJagged)
        Console.WriteLine(rows(1).Length)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["10", "10", "3"]);
}
