// vybe-test: csharp/csharp_finally_cleanup_matrix/finally_cleanup_matrix_array_length_stable
// origin: languages/csharp/tests/csharp/test_csharp_finally_cleanup_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// finally_cleanup_matrix
int seed = 54; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; __Check((numbers.Length == 3).ToString(), "True");
