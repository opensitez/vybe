// vybe-test: csharp/csharp_threading_pool_matrix/threading_pool_matrix_arithmetic_zeroing
// origin: languages/csharp/tests/csharp/test_csharp_threading_pool_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// threading_pool_matrix
int seed = 87; __Check((seed - seed == 0).ToString(), "True");
