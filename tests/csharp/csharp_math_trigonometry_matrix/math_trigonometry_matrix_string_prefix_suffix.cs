// vybe-test: csharp/csharp_math_trigonometry_matrix/math_trigonometry_matrix_string_prefix_suffix
// origin: languages/csharp/tests/csharp/test_csharp_math_trigonometry_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// math_trigonometry_matrix
string feature = "math_trigonometry_matrix"; __Check((feature.Substring(0, 1) == feature[0].ToString()).ToString(), "True");
