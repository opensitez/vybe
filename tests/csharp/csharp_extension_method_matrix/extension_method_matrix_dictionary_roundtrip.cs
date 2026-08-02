// vybe-test: csharp/csharp_extension_method_matrix/extension_method_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_extension_method_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// extension_method_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[78] = 79; __Check((map.ContainsKey(78) && map[78] == 79).ToString(), "True");
