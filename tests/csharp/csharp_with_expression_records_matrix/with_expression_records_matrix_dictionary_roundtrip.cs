// vybe-test: csharp/csharp_with_expression_records_matrix/with_expression_records_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// with_expression_records_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[109] = 110; __Check((map.ContainsKey(109) && map[109] == 110).ToString(), "True");
