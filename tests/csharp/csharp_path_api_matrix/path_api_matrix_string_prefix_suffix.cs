// vybe-test: csharp/csharp_path_api_matrix/path_api_matrix_string_prefix_suffix
// origin: languages/csharp/tests/csharp/test_csharp_path_api_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// path_api_matrix
string feature = "path_api_matrix"; __Check((feature.Substring(0, 1) == feature[0].ToString()).ToString(), "True");
