// vybe-test: csharp/csharp_linq_groupby_matrix/linq_groupby_matrix_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_linq_groupby_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// linq_groupby_matrix
int? maybe = 120; __Check((maybe.HasValue && maybe.Value == 120).ToString(), "True");
