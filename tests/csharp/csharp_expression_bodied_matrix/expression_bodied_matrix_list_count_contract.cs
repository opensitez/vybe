// vybe-test: csharp/csharp_expression_bodied_matrix/expression_bodied_matrix_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// expression_bodied_matrix
var values = new System.Collections.Generic.List<int> { 106, 107, 106 }; __Check((values.Count == 3).ToString(), "True");
