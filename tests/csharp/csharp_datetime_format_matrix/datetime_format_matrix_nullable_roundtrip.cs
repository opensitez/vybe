// vybe-test: csharp/csharp_datetime_format_matrix/datetime_format_matrix_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_datetime_format_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// datetime_format_matrix
int? maybe = 96; __Check((maybe.HasValue && maybe.Value == 96).ToString(), "True");
