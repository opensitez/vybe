// vybe-test: csharp/csharp_nullable_pattern_matching_matrix/nullable_pattern_matching_matrix_string_contains_probe
// origin: languages/csharp/tests/csharp/test_csharp_nullable_pattern_matching_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// nullable_pattern_matching_matrix
string feature = "nullable_pattern_matching_matrix"; __Check((feature.Contains("a") || !feature.Contains("a")).ToString(), "True");
