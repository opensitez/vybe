// vybe-test: csharp/csharp_math_minmax_matrix/math_minmax_matrix_string_prefix_suffix
// origin: languages/csharp/tests/csharp/test_csharp_math_minmax_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// math_minmax_matrix
string feature = "math_minmax_matrix"; __Check((feature.Substring(0, 1) == feature[0].ToString()).ToString(), "True");
