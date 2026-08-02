// vybe-test: csharp/csharp_generic_constraints_matrix/generic_constraints_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_generic_constraints_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// generic_constraints_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[80] = 81; __Check((map.ContainsKey(80) && map[80] == 81).ToString(), "True");
