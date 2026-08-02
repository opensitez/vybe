// vybe-test: csharp/csharp_regex_pattern_matrix/regex_pattern_matrix_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_regex_pattern_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// regex_pattern_matrix
int? maybe = null; int fallback = maybe ?? 99; __Check((fallback == 99).ToString(), "True");
