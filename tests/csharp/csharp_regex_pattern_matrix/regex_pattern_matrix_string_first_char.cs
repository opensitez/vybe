// vybe-test: csharp/csharp_regex_pattern_matrix/regex_pattern_matrix_string_first_char
// origin: languages/csharp/tests/csharp/test_csharp_regex_pattern_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// regex_pattern_matrix
string feature = "regex_pattern_matrix"; __Check((feature[0] == feature[0]).ToString(), "True");
