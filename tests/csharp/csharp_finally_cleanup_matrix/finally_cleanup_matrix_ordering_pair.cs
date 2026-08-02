// vybe-test: csharp/csharp_finally_cleanup_matrix/finally_cleanup_matrix_ordering_pair
// origin: languages/csharp/tests/csharp/test_csharp_finally_cleanup_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// finally_cleanup_matrix
int seed = 54; int right = seed + 1; __Check((seed < right).ToString(), "True");
