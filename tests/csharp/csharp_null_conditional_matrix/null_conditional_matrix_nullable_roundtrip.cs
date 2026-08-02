// vybe-test: csharp/csharp_null_conditional_matrix/null_conditional_matrix_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_null_conditional_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// null_conditional_matrix
int? maybe = 55; __Check((maybe.HasValue && maybe.Value == 55).ToString(), "True");
