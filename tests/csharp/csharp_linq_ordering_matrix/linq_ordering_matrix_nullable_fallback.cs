// vybe-test: csharp/csharp_linq_ordering_matrix/linq_ordering_matrix_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_linq_ordering_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// linq_ordering_matrix
int? maybe = null; int fallback = maybe ?? 121; __Check((fallback == 121).ToString(), "True");
