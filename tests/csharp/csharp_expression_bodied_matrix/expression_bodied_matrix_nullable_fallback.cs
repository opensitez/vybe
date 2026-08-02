// vybe-test: csharp/csharp_expression_bodied_matrix/expression_bodied_matrix_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// expression_bodied_matrix
int? maybe = null; int fallback = maybe ?? 106; __Check((fallback == 106).ToString(), "True");
