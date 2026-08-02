// vybe-test: csharp/csharp_extension_method_matrix/extension_method_matrix_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_extension_method_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// extension_method_matrix
double seed = 78; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
