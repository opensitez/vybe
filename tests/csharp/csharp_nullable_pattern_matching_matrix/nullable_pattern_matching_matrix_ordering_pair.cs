// vybe-test: csharp/csharp_nullable_pattern_matching_matrix/nullable_pattern_matching_matrix_ordering_pair
// origin: languages/csharp/tests/csharp/test_csharp_nullable_pattern_matching_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// nullable_pattern_matching_matrix
int seed = 125; int right = seed + 1; __Check((seed < right).ToString(), "True");
