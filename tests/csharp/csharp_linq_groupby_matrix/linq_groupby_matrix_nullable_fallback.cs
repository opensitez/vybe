// vybe-test: csharp/csharp_linq_groupby_matrix/linq_groupby_matrix_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_linq_groupby_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// linq_groupby_matrix
int? maybe = null; int fallback = maybe ?? 120; __Check((fallback == 120).ToString(), "True");
