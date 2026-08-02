// vybe-test: csharp/csharp_io_path_matrix/io_path_matrix_ordering_pair
// origin: languages/csharp/tests/csharp/test_csharp_io_path_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// io_path_matrix
int seed = 122; int right = seed + 1; __Check((seed < right).ToString(), "True");
