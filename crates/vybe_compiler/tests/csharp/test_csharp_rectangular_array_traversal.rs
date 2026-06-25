//! Multidimensional `T[,]` layout: lengths, element access, and row-major sums.
use super::helpers::run_csharp;

#[test]
fn rectangular_array_constructor_sets_both_dimensions() {
    assert_eq!(
        run_csharp(
            r#"
var grid = new int[2, 3];
Console.WriteLine(grid.GetLength(0));
Console.WriteLine(grid.GetLength(1));
"#
        ),
        &["2", "3"]
    );
}

#[test]
fn rectangular_array_initializer_literal_fills_row_major_values() {
    assert_eq!(
        run_csharp(
            r#"
int[,] grid = {
    { 1, 2 },
    { 3, 4 }
};
Console.WriteLine(grid[0, 1]);
Console.WriteLine(grid[1, 0]);
"#
        ),
        &["2", "3"]
    );
}

#[test]
fn nested_loops_sum_all_cells_in_rectangular_array() {
    assert_eq!(
        run_csharp(
            r#"
int[,] grid = {
    { 1, 2, 3 },
    { 4, 5, 6 }
};
int sum = 0;
for (int row = 0; row < grid.GetLength(0); row++) {
    for (int col = 0; col < grid.GetLength(1); col++) {
        sum += grid[row, col];
    }
}
Console.WriteLine(sum);
"#
        ),
        &["21"]
    );
}

#[test]
fn assigning_one_cell_does_not_mutate_other_rows() {
    assert_eq!(
        run_csharp(
            r#"
int[,] grid = {
    { 10, 20 },
    { 30, 40 }
};
grid[1, 1] = 99;
Console.WriteLine(grid[0, 1]);
Console.WriteLine(grid[1, 1]);
"#
        ),
        &["20", "99"]
    );
}
