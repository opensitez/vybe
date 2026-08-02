// vybe-test: csharp/csharp_finally_cleanup_matrix/finally_cleanup_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_finally_cleanup_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// finally_cleanup_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[54] = 55; __Check((map.ContainsKey(54) && map[54] == 55).ToString(), "True");
