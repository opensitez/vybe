// vybe-test: csharp/csharp_rectangular_array_traversal/rectangular_array_initializer_literal_fills_row_major_values
// origin: languages/csharp/tests/csharp/test_csharp_rectangular_array_traversal.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[,] grid = {
    { 1, 2 },
    { 3, 4 }
};
__Check((grid[0, 1]).ToString(), "2");
__Check((grid[1, 0]).ToString(), "3");
