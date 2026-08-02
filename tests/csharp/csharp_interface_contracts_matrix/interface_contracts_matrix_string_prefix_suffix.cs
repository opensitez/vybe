// vybe-test: csharp/csharp_interface_contracts_matrix/interface_contracts_matrix_string_prefix_suffix
// origin: languages/csharp/tests/csharp/test_csharp_interface_contracts_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// interface_contracts_matrix
string feature = "interface_contracts_matrix"; __Check((feature.Substring(0, 1) == feature[0].ToString()).ToString(), "True");
