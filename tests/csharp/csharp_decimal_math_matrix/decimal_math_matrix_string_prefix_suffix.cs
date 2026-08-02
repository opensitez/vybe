// vybe-test: csharp/csharp_decimal_math_matrix/decimal_math_matrix_string_prefix_suffix
// origin: languages/csharp/tests/csharp/test_csharp_decimal_math_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// decimal_math_matrix
string feature = "decimal_math_matrix"; __Check((feature.Substring(0, 1) == feature[0].ToString()).ToString(), "True");
