// vybe-test: csharp/csharp_cast_runtime_matrix/cast_runtime_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_cast_runtime_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// cast_runtime_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[61] = 62; __Check((map.ContainsKey(61) && map[61] == 62).ToString(), "True");
