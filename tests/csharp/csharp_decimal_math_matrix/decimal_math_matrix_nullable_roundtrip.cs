// vybe-test: csharp/csharp_decimal_math_matrix/decimal_math_matrix_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_decimal_math_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// decimal_math_matrix
int? maybe = 17; __Check((maybe.HasValue && maybe.Value == 17).ToString(), "True");
