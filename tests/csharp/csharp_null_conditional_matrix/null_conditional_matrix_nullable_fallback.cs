// vybe-test: csharp/csharp_null_conditional_matrix/null_conditional_matrix_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_null_conditional_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// null_conditional_matrix
int? maybe = null; int fallback = maybe ?? 55; __Check((fallback == 55).ToString(), "True");
