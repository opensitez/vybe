// vybe-test: csharp/csharp_property_accessor_matrix/property_accessor_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_property_accessor_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// property_accessor_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[64] = 65; __Check((map.ContainsKey(64) && map[64] == 65).ToString(), "True");
