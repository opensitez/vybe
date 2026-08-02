// vybe-test: csharp/csharp_threading_pool_matrix/threading_pool_matrix_ordering_pair
// origin: languages/csharp/tests/csharp/test_csharp_threading_pool_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// threading_pool_matrix
int seed = 87; int right = seed + 1; __Check((seed < right).ToString(), "True");
