// vybe-test: csharp/csharp_finally_cleanup_matrix/finally_cleanup_matrix_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_finally_cleanup_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// finally_cleanup_matrix
int? maybe = null; int fallback = maybe ?? 54; __Check((fallback == 54).ToString(), "True");
