// vybe-test: csharp/csharp_decimal_math_matrix/decimal_math_matrix_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_decimal_math_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// decimal_math_matrix
string feature = "decimal_math_matrix:17"; __Check((feature.Length >= 1).ToString(), "True");
