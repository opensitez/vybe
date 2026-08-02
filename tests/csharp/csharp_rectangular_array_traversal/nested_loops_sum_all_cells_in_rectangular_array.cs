// vybe-test: csharp/csharp_rectangular_array_traversal/nested_loops_sum_all_cells_in_rectangular_array
// origin: languages/csharp/tests/csharp/test_csharp_rectangular_array_traversal.rs

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
