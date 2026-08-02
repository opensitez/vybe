// vybe-test: csharp/csharp_threading_pool_matrix/threading_pool_matrix_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_threading_pool_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// threading_pool_matrix
int? maybe = 87; __Check((maybe.HasValue && maybe.Value == 87).ToString(), "True");
