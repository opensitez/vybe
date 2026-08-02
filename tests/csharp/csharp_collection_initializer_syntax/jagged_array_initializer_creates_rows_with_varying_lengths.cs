// vybe-test: csharp/csharp_collection_initializer_syntax/jagged_array_initializer_creates_rows_with_varying_lengths
// origin: languages/csharp/tests/csharp/test_csharp_collection_initializer_syntax.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[][] grid = {
    new[] { 1, 2 },
    new[] { 3, 4, 5 }
};
__Check((grid[1].Length).ToString(), "3");
__Check((grid[1][2]).ToString(), "5");
