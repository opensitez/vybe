// vybe-test: csharp/csharp_path_api_matrix/path_api_matrix_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_path_api_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// path_api_matrix
int? maybe = 123; __Check((maybe.HasValue && maybe.Value == 123).ToString(), "True");
