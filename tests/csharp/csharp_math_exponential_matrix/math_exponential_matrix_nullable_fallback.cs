// vybe-test: csharp/csharp_math_exponential_matrix/math_exponential_matrix_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_math_exponential_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// math_exponential_matrix
int? maybe = null; int fallback = maybe ?? 103; __Check((fallback == 103).ToString(), "True");
