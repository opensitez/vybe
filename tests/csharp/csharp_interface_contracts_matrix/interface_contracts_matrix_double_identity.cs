// vybe-test: csharp/csharp_interface_contracts_matrix/interface_contracts_matrix_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_interface_contracts_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// interface_contracts_matrix
double seed = 73; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
