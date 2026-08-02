// vybe-test: csharp/csharp_threading_pool_matrix/threading_pool_matrix_string_non_empty
// origin: languages/csharp/tests/csharp/test_csharp_threading_pool_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// threading_pool_matrix
string feature = "threading_pool_matrix"; __Check((feature.Length > 0).ToString(), "True");
