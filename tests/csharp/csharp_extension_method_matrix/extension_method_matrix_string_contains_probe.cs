// vybe-test: csharp/csharp_extension_method_matrix/extension_method_matrix_string_contains_probe
// origin: languages/csharp/tests/csharp/test_csharp_extension_method_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// extension_method_matrix
string feature = "extension_method_matrix"; __Check((feature.Contains("a") || !feature.Contains("a")).ToString(), "True");
