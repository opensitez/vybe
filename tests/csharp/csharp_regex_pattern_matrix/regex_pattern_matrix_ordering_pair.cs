// vybe-test: csharp/csharp_regex_pattern_matrix/regex_pattern_matrix_ordering_pair
// origin: languages/csharp/tests/csharp/test_csharp_regex_pattern_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// regex_pattern_matrix
int seed = 99; int right = seed + 1; __Check((seed < right).ToString(), "True");
