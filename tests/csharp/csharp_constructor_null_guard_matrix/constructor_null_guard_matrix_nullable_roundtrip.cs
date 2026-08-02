// vybe-test: csharp/csharp_constructor_null_guard_matrix/constructor_null_guard_matrix_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_constructor_null_guard_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// constructor_null_guard_matrix
int? maybe = 126; __Check((maybe.HasValue && maybe.Value == 126).ToString(), "True");
