// vybe-test: csharp/csharp_regex_pattern_matrix/regex_pattern_matrix_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_regex_pattern_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// regex_pattern_matrix
double seed = 99; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
