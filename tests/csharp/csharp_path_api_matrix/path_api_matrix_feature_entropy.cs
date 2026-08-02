// vybe-test: csharp/csharp_path_api_matrix/path_api_matrix_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_path_api_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// path_api_matrix
string feature = "path_api_matrix:123"; __Check((feature.Length >= 1).ToString(), "True");
