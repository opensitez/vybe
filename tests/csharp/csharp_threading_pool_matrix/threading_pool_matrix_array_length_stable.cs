// vybe-test: csharp/csharp_threading_pool_matrix/threading_pool_matrix_array_length_stable
// origin: languages/csharp/tests/csharp/test_csharp_threading_pool_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// threading_pool_matrix
int seed = 87; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; __Check((numbers.Length == 3).ToString(), "True");
