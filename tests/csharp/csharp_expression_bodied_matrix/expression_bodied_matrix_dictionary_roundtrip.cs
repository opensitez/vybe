// vybe-test: csharp/csharp_expression_bodied_matrix/expression_bodied_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// expression_bodied_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[106] = 107; __Check((map.ContainsKey(106) && map[106] == 107).ToString(), "True");
