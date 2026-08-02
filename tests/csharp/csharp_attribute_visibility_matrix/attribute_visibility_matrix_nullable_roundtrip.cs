// vybe-test: csharp/csharp_attribute_visibility_matrix/attribute_visibility_matrix_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_attribute_visibility_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// attribute_visibility_matrix
int? maybe = 93; __Check((maybe.HasValue && maybe.Value == 93).ToString(), "True");
