// vybe-test: csharp/csharp_linq_groupby_matrix/linq_groupby_matrix_array_length_stable
// origin: languages/csharp/tests/csharp/test_csharp_linq_groupby_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// linq_groupby_matrix
int seed = 120; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; __Check((numbers.Length == 3).ToString(), "True");
