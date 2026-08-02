// vybe-test: csharp/csharp_constructor_chaining_matrix/constructor_chaining_matrix_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_constructor_chaining_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// constructor_chaining_matrix
int? maybe = 68; __Check((maybe.HasValue && maybe.Value == 68).ToString(), "True");
