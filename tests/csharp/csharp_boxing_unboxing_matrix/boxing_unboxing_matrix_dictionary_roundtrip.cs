// vybe-test: csharp/csharp_boxing_unboxing_matrix/boxing_unboxing_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_boxing_unboxing_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// boxing_unboxing_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[62] = 63; __Check((map.ContainsKey(62) && map[62] == 63).ToString(), "True");
