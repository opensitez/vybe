// vybe-test: csharp/csharp_interpolation_verbatim_matrix/interpolation_verbatim_matrix_ordering_pair
// origin: languages/csharp/tests/csharp/test_csharp_interpolation_verbatim_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// interpolation_verbatim_matrix
int seed = 110; int right = seed + 1; __Check((seed < right).ToString(), "True");
