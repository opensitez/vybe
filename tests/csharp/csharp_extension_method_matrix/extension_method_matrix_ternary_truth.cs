// vybe-test: csharp/csharp_extension_method_matrix/extension_method_matrix_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_extension_method_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// extension_method_matrix
int seed = 78; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
