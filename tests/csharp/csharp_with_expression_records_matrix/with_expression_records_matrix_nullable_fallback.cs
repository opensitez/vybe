// vybe-test: csharp/csharp_with_expression_records_matrix/with_expression_records_matrix_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// with_expression_records_matrix
int? maybe = null; int fallback = maybe ?? 109; __Check((fallback == 109).ToString(), "True");
