// vybe-test: csharp/csharp_property_accessor_matrix/property_accessor_matrix_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_property_accessor_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// property_accessor_matrix
int? maybe = 64; __Check((maybe.HasValue && maybe.Value == 64).ToString(), "True");
