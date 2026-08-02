// vybe-test: csharp/csharp_path_api_matrix/path_api_matrix_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_path_api_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// path_api_matrix
double seed = 123; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
