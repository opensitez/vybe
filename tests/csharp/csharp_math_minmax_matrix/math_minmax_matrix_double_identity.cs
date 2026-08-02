// vybe-test: csharp/csharp_math_minmax_matrix/math_minmax_matrix_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_math_minmax_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// math_minmax_matrix
double seed = 101; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
