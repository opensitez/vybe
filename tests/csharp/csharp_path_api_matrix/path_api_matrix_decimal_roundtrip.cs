// vybe-test: csharp/csharp_path_api_matrix/path_api_matrix_decimal_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_path_api_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// path_api_matrix
decimal amount = 10m; __Check(((amount / 2m) * 2m == 10m).ToString(), "True");
