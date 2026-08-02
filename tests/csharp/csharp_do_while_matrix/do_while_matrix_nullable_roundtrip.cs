// vybe-test: csharp/csharp_do_while_matrix/do_while_matrix_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_do_while_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// do_while_matrix
int? maybe = 48; __Check((maybe.HasValue && maybe.Value == 48).ToString(), "True");
