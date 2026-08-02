// vybe-test: csharp/csharp_random_distribution_matrix/random_distribution_matrix_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_random_distribution_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// random_distribution_matrix
int? maybe = null; int fallback = maybe ?? 98; __Check((fallback == 98).ToString(), "True");
