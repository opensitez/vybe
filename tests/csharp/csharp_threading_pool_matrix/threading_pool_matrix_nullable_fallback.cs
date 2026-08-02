// vybe-test: csharp/csharp_threading_pool_matrix/threading_pool_matrix_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_threading_pool_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// threading_pool_matrix
int? maybe = null; int fallback = maybe ?? 87; __Check((fallback == 87).ToString(), "True");
