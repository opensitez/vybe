// vybe-test: csharp/csharp_generic_variance_matrix/generic_variance_matrix_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_generic_variance_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// generic_variance_matrix
int? maybe = null; int fallback = maybe ?? 82; __Check((fallback == 82).ToString(), "True");
