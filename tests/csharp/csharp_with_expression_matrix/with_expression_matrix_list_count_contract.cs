// vybe-test: csharp/csharp_with_expression_matrix/with_expression_matrix_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// with_expression_matrix
var values = new System.Collections.Generic.List<int> { 108, 109, 108 }; __Check((values.Count == 3).ToString(), "True");
