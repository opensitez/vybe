// vybe-test: csharp/csharp_regex_pattern_matrix/regex_pattern_matrix_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_regex_pattern_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// regex_pattern_matrix
int? maybe = 99; __Check((maybe.HasValue && maybe.Value == 99).ToString(), "True");
