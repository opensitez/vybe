// vybe-test: csharp/csharp_linq_join_matrix/linq_join_matrix_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_linq_join_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// linq_join_matrix
int? maybe = 119; __Check((maybe.HasValue && maybe.Value == 119).ToString(), "True");
