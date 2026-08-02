// vybe-test: csharp/csharp_math_minmax_matrix/math_minmax_matrix_ordering_pair
// origin: languages/csharp/tests/csharp/test_csharp_math_minmax_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// math_minmax_matrix
int seed = 101; int right = seed + 1; __Check((seed < right).ToString(), "True");
