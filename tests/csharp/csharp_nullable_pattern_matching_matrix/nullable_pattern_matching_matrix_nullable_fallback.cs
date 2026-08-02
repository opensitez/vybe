// vybe-test: csharp/csharp_nullable_pattern_matching_matrix/nullable_pattern_matching_matrix_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_nullable_pattern_matching_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// nullable_pattern_matching_matrix
int? maybe = null; int fallback = maybe ?? 125; __Check((fallback == 125).ToString(), "True");
