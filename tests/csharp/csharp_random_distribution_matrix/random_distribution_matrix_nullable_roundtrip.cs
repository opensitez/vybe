// vybe-test: csharp/csharp_random_distribution_matrix/random_distribution_matrix_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_random_distribution_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// random_distribution_matrix
int? maybe = 98; __Check((maybe.HasValue && maybe.Value == 98).ToString(), "True");
