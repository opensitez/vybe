// vybe-test: csharp/csharp_threading_pool_matrix/threading_pool_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_threading_pool_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// threading_pool_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[87] = 88; __Check((map.ContainsKey(87) && map[87] == 88).ToString(), "True");
