// vybe-test: csharp/csharp_path_api_matrix/path_api_matrix_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_path_api_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// path_api_matrix
int? maybe = null; int fallback = maybe ?? 123; __Check((fallback == 123).ToString(), "True");
