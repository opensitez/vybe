// vybe-test: csharp/csharp_math_exponential_matrix/math_exponential_matrix_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_math_exponential_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// math_exponential_matrix
int? maybe = 103; __Check((maybe.HasValue && maybe.Value == 103).ToString(), "True");
