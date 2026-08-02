// vybe-test: csharp/csharp_interpolation_basic_matrix/interpolation_basic_matrix_string_non_empty
// origin: languages/csharp/tests/csharp/test_csharp_interpolation_basic_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// interpolation_basic_matrix
string feature = "interpolation_basic_matrix"; __Check((feature.Length > 0).ToString(), "True");
