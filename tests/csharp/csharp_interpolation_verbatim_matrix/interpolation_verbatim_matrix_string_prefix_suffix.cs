// vybe-test: csharp/csharp_interpolation_verbatim_matrix/interpolation_verbatim_matrix_string_prefix_suffix
// origin: languages/csharp/tests/csharp/test_csharp_interpolation_verbatim_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// interpolation_verbatim_matrix
string feature = "interpolation_verbatim_matrix"; __Check((feature.Substring(0, 1) == feature[0].ToString()).ToString(), "True");
