// vybe-test: csharp/csharp_extension_method_matrix/extension_method_matrix_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_extension_method_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// extension_method_matrix
string feature = "extension_method_matrix:78"; __Check((feature.Length >= 1).ToString(), "True");
