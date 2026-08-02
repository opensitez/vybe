// vybe-test: csharp/csharp_with_expression_matrix/with_expression_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// with_expression_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[108] = 109; __Check((map.ContainsKey(108) && map[108] == 109).ToString(), "True");
