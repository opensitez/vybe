// vybe-test: csharp/csharp_null_conditional_matrix/null_conditional_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_null_conditional_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// null_conditional_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[55] = 56; __Check((map.ContainsKey(55) && map[55] == 56).ToString(), "True");
