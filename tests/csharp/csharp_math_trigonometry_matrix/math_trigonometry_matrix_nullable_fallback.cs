// vybe-test: csharp/csharp_math_trigonometry_matrix/math_trigonometry_matrix_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_math_trigonometry_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// math_trigonometry_matrix
int? maybe = null; int fallback = maybe ?? 102; __Check((fallback == 102).ToString(), "True");
