//! Jagged arrays, multi-dimensional arrays, and array rank inspection.
use super::helpers::run_csharp;

#[test]
fn jagged_array_rows_have_independent_lengths() {
    assert_eq!(
        run_csharp(
            r#"int[][] jag = new int[3][];
jag[0] = new int[]{1};
jag[1] = new int[]{2,3};
jag[2] = new int[]{4,5,6};
Console.WriteLine(jag[2].Length);"#
        ),
        &["3"]
    );
}

#[test]
fn jagged_array_element_access_uses_double_indexer() {
    assert_eq!(
        run_csharp(
            r#"int[][] jag = new int[][]{ new[]{10,20}, new[]{30,40,50} };
Console.WriteLine(jag[1][2]);"#
        ),
        &["50"]
    );
}

#[test]
fn two_dimensional_array_get_length_returns_dimension_size() {
    assert_eq!(
        run_csharp(
            r#"int[,] grid = new int[3,4];
Console.WriteLine(grid.GetLength(0));
Console.WriteLine(grid.GetLength(1));"#
        ),
        &["3", "4"]
    );
}

#[test]
fn two_dimensional_array_element_set_and_read() {
    assert_eq!(
        run_csharp(
            r#"int[,] m = new int[2,2];
m[0,1] = 7;
Console.WriteLine(m[0,1]);"#
        ),
        &["7"]
    );
}

#[test]
fn array_rank_is_two_for_2d_array() {
    assert_eq!(
        run_csharp(r#"int[,] a = new int[2,3]; Console.WriteLine(a.Rank);"#),
        &["2"]
    );
}

#[test]
fn array_rank_is_one_for_flat_array() {
    assert_eq!(
        run_csharp(r#"int[] a = new int[5]; Console.WriteLine(a.Rank);"#),
        &["1"]
    );
}

#[test]
fn jagged_array_foreach_over_rows_sums_correctly() {
    assert_eq!(
        run_csharp(
            r#"int[][] jag = new[]{ new[]{1,2}, new[]{3,4,5} };
int total=0;
foreach(var row in jag) foreach(var v in row) total+=v;
Console.WriteLine(total);"#
        ),
        &["15"]
    );
}
