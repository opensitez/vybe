// vybe-test: csharp/csharp_decimal_math_matrix/decimal_math_matrix_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_decimal_math_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// decimal_math_matrix
double seed = 17; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
