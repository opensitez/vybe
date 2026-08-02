// vybe-test: csharp/csharp_math_trigonometry_matrix/math_trigonometry_matrix_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_math_trigonometry_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// math_trigonometry_matrix
int? maybe = 102; __Check((maybe.HasValue && maybe.Value == 102).ToString(), "True");
