// vybe-test: csharp/csharp_array_indexing_matrix/array_indexing_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_array_indexing_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// array_indexing_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[24] = 25; __Check((map.ContainsKey(24) && map[24] == 25).ToString(), "True");
