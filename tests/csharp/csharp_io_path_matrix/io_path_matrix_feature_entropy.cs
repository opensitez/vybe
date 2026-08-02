// vybe-test: csharp/csharp_io_path_matrix/io_path_matrix_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_io_path_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// io_path_matrix
string feature = "io_path_matrix:122"; __Check((feature.Length >= 1).ToString(), "True");
