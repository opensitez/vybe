// vybe-test: csharp/csharp_constructor_chaining_matrix/constructor_chaining_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_constructor_chaining_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// constructor_chaining_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[68] = 69; __Check((map.ContainsKey(68) && map[68] == 69).ToString(), "True");
