// vybe-test: csharp/csharp_linq_join_matrix/linq_join_matrix_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_linq_join_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// linq_join_matrix
int? maybe = null; int fallback = maybe ?? 119; __Check((fallback == 119).ToString(), "True");
