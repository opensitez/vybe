// vybe-test: csharp/csharp_cast_runtime_matrix/cast_runtime_matrix_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_cast_runtime_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// cast_runtime_matrix
int? maybe = null; int fallback = maybe ?? 61; __Check((fallback == 61).ToString(), "True");
