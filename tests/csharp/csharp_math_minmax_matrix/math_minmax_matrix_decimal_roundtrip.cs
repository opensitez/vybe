// vybe-test: csharp/csharp_math_minmax_matrix/math_minmax_matrix_decimal_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_math_minmax_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// math_minmax_matrix
decimal amount = 10m; __Check(((amount / 2m) * 2m == 10m).ToString(), "True");
