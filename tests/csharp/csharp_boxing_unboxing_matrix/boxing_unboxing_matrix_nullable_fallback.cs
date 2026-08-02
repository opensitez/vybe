// vybe-test: csharp/csharp_boxing_unboxing_matrix/boxing_unboxing_matrix_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_boxing_unboxing_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// boxing_unboxing_matrix
int? maybe = null; int fallback = maybe ?? 62; __Check((fallback == 62).ToString(), "True");
