// vybe-test: csharp/csharp_interpolation_basic_matrix/interpolation_basic_matrix_arithmetic_increment
// origin: languages/csharp/tests/csharp/test_csharp_interpolation_basic_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// interpolation_basic_matrix
int seed = 112; __Check((seed + 1 > seed).ToString(), "True");
