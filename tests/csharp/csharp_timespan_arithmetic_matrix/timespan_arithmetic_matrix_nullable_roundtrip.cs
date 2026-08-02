// vybe-test: csharp/csharp_timespan_arithmetic_matrix/timespan_arithmetic_matrix_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_timespan_arithmetic_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// timespan_arithmetic_matrix
int? maybe = 95; __Check((maybe.HasValue && maybe.Value == 95).ToString(), "True");
