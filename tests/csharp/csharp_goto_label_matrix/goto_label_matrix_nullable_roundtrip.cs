// vybe-test: csharp/csharp_goto_label_matrix/goto_label_matrix_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_goto_label_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// goto_label_matrix
int? maybe = 50; __Check((maybe.HasValue && maybe.Value == 50).ToString(), "True");
