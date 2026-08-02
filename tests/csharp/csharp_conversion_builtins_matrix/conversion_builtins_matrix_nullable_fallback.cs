// vybe-test: csharp/csharp_conversion_builtins_matrix/conversion_builtins_matrix_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_conversion_builtins_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// conversion_builtins_matrix
int? maybe = null; int fallback = maybe ?? 124; __Check((fallback == 124).ToString(), "True");
