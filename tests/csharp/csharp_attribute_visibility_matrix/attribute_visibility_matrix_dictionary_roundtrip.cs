// vybe-test: csharp/csharp_attribute_visibility_matrix/attribute_visibility_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_attribute_visibility_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// attribute_visibility_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[93] = 94; __Check((map.ContainsKey(93) && map[93] == 94).ToString(), "True");
