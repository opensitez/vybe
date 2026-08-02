// vybe-test: csharp/csharp_math_trigonometry_matrix/math_trigonometry_matrix_string_first_char
// origin: languages/csharp/tests/csharp/test_csharp_math_trigonometry_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// math_trigonometry_matrix
string feature = "math_trigonometry_matrix"; __Check((feature[0] == feature[0]).ToString(), "True");
