// vybe-test: csharp/csharp_conversion_builtins_matrix/conversion_builtins_matrix_ordering_pair
// origin: languages/csharp/tests/csharp/test_csharp_conversion_builtins_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// conversion_builtins_matrix
int seed = 124; int right = seed + 1; __Check((seed < right).ToString(), "True");
