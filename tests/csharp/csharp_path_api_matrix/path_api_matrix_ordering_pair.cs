// vybe-test: csharp/csharp_path_api_matrix/path_api_matrix_ordering_pair
// origin: languages/csharp/tests/csharp/test_csharp_path_api_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// path_api_matrix
int seed = 123; int right = seed + 1; __Check((seed < right).ToString(), "True");
