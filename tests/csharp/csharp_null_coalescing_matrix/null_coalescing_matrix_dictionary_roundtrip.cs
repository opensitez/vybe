// vybe-test: csharp/csharp_null_coalescing_matrix/null_coalescing_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_null_coalescing_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// null_coalescing_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[56] = 57; __Check((map.ContainsKey(56) && map[56] == 57).ToString(), "True");
