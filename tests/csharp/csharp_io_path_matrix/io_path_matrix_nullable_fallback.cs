// vybe-test: csharp/csharp_io_path_matrix/io_path_matrix_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_io_path_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// io_path_matrix
int? maybe = null; int fallback = maybe ?? 122; __Check((fallback == 122).ToString(), "True");
