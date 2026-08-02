// vybe-test: csharp/csharp_math_trigonometry_matrix/math_trigonometry_matrix_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_math_trigonometry_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// math_trigonometry_matrix
string feature = "math_trigonometry_matrix:102"; __Check((feature.Length >= 1).ToString(), "True");
