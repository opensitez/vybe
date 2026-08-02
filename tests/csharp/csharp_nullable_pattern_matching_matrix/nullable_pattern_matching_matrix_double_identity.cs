// vybe-test: csharp/csharp_nullable_pattern_matching_matrix/nullable_pattern_matching_matrix_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_nullable_pattern_matching_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// nullable_pattern_matching_matrix
double seed = 125; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
