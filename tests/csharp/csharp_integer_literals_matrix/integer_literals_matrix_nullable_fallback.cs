// vybe-test: csharp/csharp_integer_literals_matrix/integer_literals_matrix_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_integer_literals_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// integer_literals_matrix
int? maybe = null; int fallback = maybe ?? 15; __Check((fallback == 15).ToString(), "True");
