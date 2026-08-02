// vybe-test: csharp/csharp_interface_contracts_matrix/interface_contracts_matrix_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_interface_contracts_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// interface_contracts_matrix
string feature = "interface_contracts_matrix:73"; __Check((feature.Length >= 1).ToString(), "True");
