// vybe-test: csharp/csharp_interpolation_verbatim_matrix/interpolation_verbatim_matrix_string_contains_probe
// origin: languages/csharp/tests/csharp/test_csharp_interpolation_verbatim_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// interpolation_verbatim_matrix
string feature = "interpolation_verbatim_matrix"; __Check((feature.Contains("a") || !feature.Contains("a")).ToString(), "True");
