// vybe-test: csharp/csharp_threading_pool_matrix/threading_pool_matrix_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_threading_pool_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// threading_pool_matrix
double seed = 87; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
