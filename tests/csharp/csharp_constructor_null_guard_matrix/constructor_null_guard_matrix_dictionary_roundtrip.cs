// vybe-test: csharp/csharp_constructor_null_guard_matrix/constructor_null_guard_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_constructor_null_guard_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// constructor_null_guard_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[126] = 127; __Check((map.ContainsKey(126) && map[126] == 127).ToString(), "True");
