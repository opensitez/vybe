// vybe-test: csharp/csharp_interface_contracts_matrix/interface_contracts_matrix_ordering_pair
// origin: languages/csharp/tests/csharp/test_csharp_interface_contracts_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// interface_contracts_matrix
int seed = 73; int right = seed + 1; __Check((seed < right).ToString(), "True");
