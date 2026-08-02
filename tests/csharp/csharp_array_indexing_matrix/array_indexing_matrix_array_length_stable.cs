// vybe-test: csharp/csharp_array_indexing_matrix/array_indexing_matrix_array_length_stable
// origin: languages/csharp/tests/csharp/test_csharp_array_indexing_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// array_indexing_matrix
int seed = 24; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; __Check((numbers.Length == 3).ToString(), "True");
