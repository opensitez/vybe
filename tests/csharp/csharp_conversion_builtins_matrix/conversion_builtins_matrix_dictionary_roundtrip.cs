// vybe-test: csharp/csharp_conversion_builtins_matrix/conversion_builtins_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_conversion_builtins_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// conversion_builtins_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[124] = 125; __Check((map.ContainsKey(124) && map[124] == 125).ToString(), "True");
