// vybe-test: csharp/csharp_io_path_matrix/io_path_matrix_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_io_path_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// io_path_matrix
double seed = 122; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
