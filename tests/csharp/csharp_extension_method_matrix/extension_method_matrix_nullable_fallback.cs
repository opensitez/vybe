// vybe-test: csharp/csharp_extension_method_matrix/extension_method_matrix_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_extension_method_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// extension_method_matrix
int? maybe = null; int fallback = maybe ?? 78; __Check((fallback == 78).ToString(), "True");
