// vybe-test: csharp/csharp_interpolation_basic_matrix/interpolation_basic_matrix_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_interpolation_basic_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// interpolation_basic_matrix
string feature = "interpolation_basic_matrix:112"; __Check((feature.Length >= 1).ToString(), "True");
