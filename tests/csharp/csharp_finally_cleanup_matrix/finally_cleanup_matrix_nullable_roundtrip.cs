// vybe-test: csharp/csharp_finally_cleanup_matrix/finally_cleanup_matrix_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_finally_cleanup_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// finally_cleanup_matrix
int? maybe = 54; __Check((maybe.HasValue && maybe.Value == 54).ToString(), "True");
