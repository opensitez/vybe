// vybe-test: csharp/csharp_path_api_matrix/path_api_matrix_string_contains_probe
// origin: languages/csharp/tests/csharp/test_csharp_path_api_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// path_api_matrix
string feature = "path_api_matrix"; __Check((feature.Contains("a") || !feature.Contains("a")).ToString(), "True");
