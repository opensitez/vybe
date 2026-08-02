// vybe-test: csharp/csharp_interface_contracts_matrix/interface_contracts_matrix_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_interface_contracts_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// interface_contracts_matrix
var set = new System.Collections.Generic.HashSet<int>(); set.Add(73); set.Add(73); __Check((set.Count == 1).ToString(), "True");
