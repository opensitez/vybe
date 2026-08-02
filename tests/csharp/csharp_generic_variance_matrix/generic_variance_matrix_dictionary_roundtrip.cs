// vybe-test: csharp/csharp_generic_variance_matrix/generic_variance_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_generic_variance_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// generic_variance_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[82] = 83; __Check((map.ContainsKey(82) && map[82] == 83).ToString(), "True");
