// vybe-test: csharp/csharp_null_conditional_matrix/null_conditional_matrix_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_null_conditional_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// null_conditional_matrix
var values = new System.Collections.Generic.List<int> { 55, 56, 55 }; __Check((values.Count == 3).ToString(), "True");
