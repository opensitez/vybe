// vybe-test: csharp/csharp_pointer_like_emulation_matrix/pointer_like_emulation_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_pointer_like_emulation_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// pointer_like_emulation_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[114] = 115; __Check((map.ContainsKey(114) && map[114] == 115).ToString(), "True");
