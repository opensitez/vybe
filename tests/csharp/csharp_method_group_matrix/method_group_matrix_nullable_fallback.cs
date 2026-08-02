// vybe-test: csharp/csharp_method_group_matrix/method_group_matrix_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_method_group_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// method_group_matrix
int? maybe = null; int fallback = maybe ?? 79; __Check((fallback == 79).ToString(), "True");
