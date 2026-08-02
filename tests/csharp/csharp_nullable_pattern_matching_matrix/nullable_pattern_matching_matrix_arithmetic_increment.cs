// vybe-test: csharp/csharp_nullable_pattern_matching_matrix/nullable_pattern_matching_matrix_arithmetic_increment
// origin: languages/csharp/tests/csharp/test_csharp_nullable_pattern_matching_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// nullable_pattern_matching_matrix
int seed = 125; __Check((seed + 1 > seed).ToString(), "True");
