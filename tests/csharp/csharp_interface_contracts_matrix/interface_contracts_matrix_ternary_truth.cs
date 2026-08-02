// vybe-test: csharp/csharp_interface_contracts_matrix/interface_contracts_matrix_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_interface_contracts_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// interface_contracts_matrix
int seed = 73; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
