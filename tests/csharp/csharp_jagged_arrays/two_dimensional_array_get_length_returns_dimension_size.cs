// vybe-test: csharp/csharp_jagged_arrays/two_dimensional_array_get_length_returns_dimension_size
// origin: languages/csharp/tests/csharp/test_csharp_jagged_arrays.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[,] grid = new int[3,4];
__Check((grid.GetLength(0)).ToString(), "3");
__Check((grid.GetLength(1)).ToString(), "4");
