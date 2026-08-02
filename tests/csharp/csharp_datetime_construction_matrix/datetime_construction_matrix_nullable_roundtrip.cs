// vybe-test: csharp/csharp_datetime_construction_matrix/datetime_construction_matrix_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_datetime_construction_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// datetime_construction_matrix
int? maybe = 94; __Check((maybe.HasValue && maybe.Value == 94).ToString(), "True");
