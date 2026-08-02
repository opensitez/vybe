// vybe-test: csharp/csharp_interface_contracts_matrix/interface_contracts_matrix_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_interface_contracts_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// interface_contracts_matrix
var tuple = (left: 73, right: 74); __Check((tuple.left < tuple.right).ToString(), "True");
