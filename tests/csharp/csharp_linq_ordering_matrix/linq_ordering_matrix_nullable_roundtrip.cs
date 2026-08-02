// vybe-test: csharp/csharp_linq_ordering_matrix/linq_ordering_matrix_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_linq_ordering_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// linq_ordering_matrix
int? maybe = 121; __Check((maybe.HasValue && maybe.Value == 121).ToString(), "True");
