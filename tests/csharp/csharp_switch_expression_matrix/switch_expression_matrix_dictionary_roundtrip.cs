// vybe-test: csharp/csharp_switch_expression_matrix/switch_expression_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// switch_expression_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[43] = 44; __Check((map.ContainsKey(43) && map[43] == 44).ToString(), "True");
