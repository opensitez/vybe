// vybe-test: csharp/csharp_extension_method_matrix/extension_method_matrix_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_extension_method_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// extension_method_matrix
int? maybe = 78; __Check((maybe.HasValue && maybe.Value == 78).ToString(), "True");
