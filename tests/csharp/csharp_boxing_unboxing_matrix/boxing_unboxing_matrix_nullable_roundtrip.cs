// vybe-test: csharp/csharp_boxing_unboxing_matrix/boxing_unboxing_matrix_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_boxing_unboxing_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// boxing_unboxing_matrix
int? maybe = 62; __Check((maybe.HasValue && maybe.Value == 62).ToString(), "True");
