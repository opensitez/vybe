// vybe-test: csharp/csharp_interface_contracts_matrix/interface_contracts_matrix_arithmetic_increment
// origin: languages/csharp/tests/csharp/test_csharp_interface_contracts_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// interface_contracts_matrix
int seed = 73; __Check((seed + 1 > seed).ToString(), "True");
