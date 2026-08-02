// vybe-test: csharp/csharp_goto_label_matrix/goto_label_matrix_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_goto_label_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// goto_label_matrix
int? maybe = null; int fallback = maybe ?? 50; __Check((fallback == 50).ToString(), "True");
