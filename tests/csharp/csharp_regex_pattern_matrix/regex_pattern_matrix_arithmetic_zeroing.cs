// vybe-test: csharp/csharp_regex_pattern_matrix/regex_pattern_matrix_arithmetic_zeroing
// origin: languages/csharp/tests/csharp/test_csharp_regex_pattern_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// regex_pattern_matrix
int seed = 99; __Check((seed - seed == 0).ToString(), "True");
