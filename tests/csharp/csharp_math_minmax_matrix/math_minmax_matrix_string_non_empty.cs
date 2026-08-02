// vybe-test: csharp/csharp_math_minmax_matrix/math_minmax_matrix_string_non_empty
// origin: languages/csharp/tests/csharp/test_csharp_math_minmax_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// math_minmax_matrix
string feature = "math_minmax_matrix"; __Check((feature.Length > 0).ToString(), "True");
