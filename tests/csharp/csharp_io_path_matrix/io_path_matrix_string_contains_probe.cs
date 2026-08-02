// vybe-test: csharp/csharp_io_path_matrix/io_path_matrix_string_contains_probe
// origin: languages/csharp/tests/csharp/test_csharp_io_path_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// io_path_matrix
string feature = "io_path_matrix"; __Check((feature.Contains("a") || !feature.Contains("a")).ToString(), "True");
