// vybe-test: csharp/csharp_regex_pattern_matrix/regex_pattern_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_regex_pattern_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// regex_pattern_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[99] = 100; __Check((map.ContainsKey(99) && map[99] == 100).ToString(), "True");
