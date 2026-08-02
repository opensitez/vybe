// vybe-test: csharp/csharp_async_enumerator_matrix/async_enumerator_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_async_enumerator_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// async_enumerator_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[116] = 117; __Check((map.ContainsKey(116) && map[116] == 117).ToString(), "True");
