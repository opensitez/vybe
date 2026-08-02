// vybe-test: csharp/csharp_interface_contracts_matrix/interface_contracts_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_interface_contracts_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// interface_contracts_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[73] = 74; __Check((map.ContainsKey(73) && map[73] == 74).ToString(), "True");
