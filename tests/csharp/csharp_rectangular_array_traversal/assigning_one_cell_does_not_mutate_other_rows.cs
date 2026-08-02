// vybe-test: csharp/csharp_rectangular_array_traversal/assigning_one_cell_does_not_mutate_other_rows
// origin: languages/csharp/tests/csharp/test_csharp_rectangular_array_traversal.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[,] grid = {
    { 10, 20 },
    { 30, 40 }
};
grid[1, 1] = 99;
__Check((grid[0, 1]).ToString(), "20");
__Check((grid[1, 1]).ToString(), "99");
