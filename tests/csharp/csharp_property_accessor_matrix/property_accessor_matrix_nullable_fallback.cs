// vybe-test: csharp/csharp_property_accessor_matrix/property_accessor_matrix_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_property_accessor_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// property_accessor_matrix
int? maybe = null; int fallback = maybe ?? 64; __Check((fallback == 64).ToString(), "True");
