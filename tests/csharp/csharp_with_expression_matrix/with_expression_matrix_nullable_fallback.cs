// vybe-test: csharp/csharp_with_expression_matrix/with_expression_matrix_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// with_expression_matrix
int? maybe = null; int fallback = maybe ?? 108; __Check((fallback == 108).ToString(), "True");
