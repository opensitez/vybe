// vybe-test: csharp/csharp_switch_expression_matrix/switch_expression_matrix_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// switch_expression_matrix
var values = new System.Collections.Generic.List<int> { 43, 44, 43 }; __Check((values.Count == 3).ToString(), "True");
