// vybe-test: csharp/csharp_interface_contracts_matrix/interface_contracts_matrix_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_interface_contracts_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// interface_contracts_matrix
var values = new System.Collections.Generic.List<int> { 73, 74, 73 }; __Check((values.Count == 3).ToString(), "True");
