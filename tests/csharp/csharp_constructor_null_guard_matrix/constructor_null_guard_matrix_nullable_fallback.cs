// vybe-test: csharp/csharp_constructor_null_guard_matrix/constructor_null_guard_matrix_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_constructor_null_guard_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// constructor_null_guard_matrix
int? maybe = null; int fallback = maybe ?? 126; __Check((fallback == 126).ToString(), "True");
