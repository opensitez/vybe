// vybe-test: csharp/csharp_interpolation_basic_matrix/interpolation_basic_matrix_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_interpolation_basic_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// interpolation_basic_matrix
double seed = 112; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
