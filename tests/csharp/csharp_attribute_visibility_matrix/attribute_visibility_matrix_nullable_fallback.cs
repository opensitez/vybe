// vybe-test: csharp/csharp_attribute_visibility_matrix/attribute_visibility_matrix_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_attribute_visibility_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// attribute_visibility_matrix
int? maybe = null; int fallback = maybe ?? 93; __Check((fallback == 93).ToString(), "True");
