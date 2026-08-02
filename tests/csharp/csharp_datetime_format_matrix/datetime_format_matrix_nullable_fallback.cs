// vybe-test: csharp/csharp_datetime_format_matrix/datetime_format_matrix_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_datetime_format_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// datetime_format_matrix
int? maybe = null; int fallback = maybe ?? 96; __Check((fallback == 96).ToString(), "True");
