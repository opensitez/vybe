// vybe-test: csharp/csharp_math_trigonometry_matrix/math_trigonometry_matrix_ordering_pair
// origin: languages/csharp/tests/csharp/test_csharp_math_trigonometry_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// math_trigonometry_matrix
int seed = 102; int right = seed + 1; __Check((seed < right).ToString(), "True");
