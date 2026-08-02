// vybe-test: csharp/csharp_switch_expression_matrix/switch_expression_matrix_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// switch_expression_matrix
int? maybe = null; int fallback = maybe ?? 43; __Check((fallback == 43).ToString(), "True");
