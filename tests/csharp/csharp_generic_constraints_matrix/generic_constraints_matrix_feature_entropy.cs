// vybe-test: csharp/csharp_generic_constraints_matrix/generic_constraints_matrix_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_generic_constraints_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// generic_constraints_matrix
string feature = "generic_constraints_matrix:80"; __Check((feature.Length >= 1).ToString(), "True");
