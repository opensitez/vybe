// vybe-test: csharp/csharp_null_coalescing_matrix/null_coalescing_matrix_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_null_coalescing_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// null_coalescing_matrix
int? maybe = null; int fallback = maybe ?? 56; __Check((fallback == 56).ToString(), "True");
